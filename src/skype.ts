/*
Copyright 2020 mx-puppet-skype
Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at
    http://www.apache.org/licenses/LICENSE-2.0
Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/

import {
	PuppetBridge, IRemoteUser, IRemoteRoom, IReceiveParams, IMessageEvent, IFileEvent, Log, MessageDeduplicator, Util,
	ExpireSet, IRetList,
} from "mx-puppet-bridge";
import { Client } from "./client";
import * as skypeHttp from "skype-http";
import { Contact as SkypeContact } from "skype-http/dist/lib/types/contact";
import { NewMediaMessage as SkypeNewMediaMessage } from "skype-http/dist/lib/interfaces/api/api";
import { UnexpectedHttpStatusError } from "skype-http/dist/lib/errors";
import * as decodeHtml from "decode-html";
import * as escapeHtml from "escape-html";
import { MatrixMessageParser } from "./matrixmessageparser";
import { SkypeMessageParser } from "./skypemessageparser";
import * as cheerio from "cheerio";

const log = new Log("SkypePuppet:skype");

const ROOM_TYPE_DM = 8;

interface ISkypePuppet {
	client: Client;
	data: any;
	deletedMessages: ExpireSet<string>;
}

interface ISkypePuppets {
	[puppetId: number]: ISkypePuppet;
}

export class Skype {
	private puppets: ISkypePuppets = {};
	private messageDeduplicator: MessageDeduplicator;
	private matrixMessageParser: MatrixMessageParser;
	private skypeMessageParser: SkypeMessageParser;
	constructor(
		private puppet: PuppetBridge,
	) {
		this.messageDeduplicator = new MessageDeduplicator();
		this.matrixMessageParser = new MatrixMessageParser();
		this.skypeMessageParser = new SkypeMessageParser();
	}

	public getUserParams(puppetId: number, contact: SkypeContact): IRemoteUser {
		return {
			puppetId,
			userId: contact.mri,
			name: contact.displayName,
			avatarUrl: contact.profile ? contact.profile.avatarUrl : null,
		};
	}

	public getRoomParams(puppetId: number, conversation: skypeHttp.Conversation): IRemoteRoom {
		const roomType = Number(conversation.id.split(":")[0]);
		const isDirect = roomType === ROOM_TYPE_DM;
		if (isDirect) {
			return {
				puppetId,
				roomId: `dm-${puppetId}-${conversation.id}`,
				isDirect: true,
			};
		}
		let avatarUrl: string | null = null;
		let name: string | null = null;
		if (conversation.threadProperties) {
			name = conversation.threadProperties.topic || null;
			if (name) {
				name = decodeHtml(name);
			}
			const picture = conversation.threadProperties.picture;
			if (picture && picture.startsWith("URL@")) {
				avatarUrl = picture.slice("URL@".length);
			}
		}
		return {
			puppetId,
			roomId: conversation.id,
			name,
			avatarUrl,
		};
	}

	public async getSendParams(puppetId: number, resource: skypeHttp.resources.Resource): Promise<IReceiveParams | null> {
		const roomType = Number(resource.conversation.split(":")[0]);
		const p = this.puppets[puppetId];
		const contact = await p.client.getContact(resource.from.raw);
		const conversation = await p.client.getConversation({
			puppetId,
			roomId: resource.conversation,
		});
		if (!contact || !conversation) {
			return null;
		}
		return {
			user: this.getUserParams(puppetId, contact),
			room: this.getRoomParams(puppetId, conversation),
			eventId: resource.id, // tslint:disable-line no-any
		};
	}

	public async stopClient(puppetId: number) {
		const p = this.puppets[puppetId];
		if (!p) {
			return;
		}
		await p.client.disconnect();
	}

	public async startClient(puppetId: number) {
		const p = this.puppets[puppetId];
		if (!p) {
			return;
		}
		p.client = new Client(p.data.username, p.data.password, p.data.state);
		const client = p.client;
		client.on("text", async (resource: skypeHttp.resources.TextResource) => {
			try {
				await this.handleSkypeText(puppetId, resource);
			} catch (err) {
				log.error("Error while handling text event", err);
			}
		});
		client.on("edit", async (resource: skypeHttp.resources.RichTextResource) => {
			try {
				await this.handleSkypeEdit(puppetId, resource);
			} catch (err) {
				log.error("Error while handling edit event", err);
			}
		});
		client.on("location", async (resource: skypeHttp.resources.RichTextLocationResource) => {
			try {

			} catch (err) {
				log.error("Error while handling location event", err);
			}
		});
		client.on("file", async (resource: skypeHttp.resources.FileResource) => {
			try {
				await this.handleSkypeFile(puppetId, resource);
			} catch (err) {
				log.error("Error while handling file event", err);
			}
		});
		client.on("typing", async (resource: skypeHttp.resources.Resource, typing: boolean) => {
			try {
				await this.handleSkypeTyping(puppetId, resource, typing);
			} catch (err) {
				log.error("Error while handling typing event", err);
			}
		});
		client.on("presence", async (resource: skypeHttp.resources.Resource) => {
			try {
				await this.handleSkypePresence(puppetId, resource);
			} catch (err) {
				log.error("Error while handling presence event", err);
			}
		});
		client.on("updateContact", async (oldContact: SkypeContact | null, newContact: SkypeContact) => {
			try {
				let update = oldContact === null;
				const newUser = this.getUserParams(puppetId, newContact);
				if (oldContact) {
					const oldUser = this.getUserParams(puppetId, oldContact);
					update = oldUser.name !== newUser.name || oldUser.avatarUrl !== newUser.avatarUrl;
				}
				if (update) {
					await this.puppet.updateUser(newUser);
				}
			} catch (err) {
				log.error("Error while handling updateContact event", err);
			}
		});
		const MINUTE = 60000;
		client.on("error", async (err: Error) => {
			const causeName = (err as any).cause ? (err as any).cause.name : "";
			log.error("Error when polling");
			log.error("name: ", err.name);
			const errr = err as any;
			if (errr.cause) {
				log.error("cause name: ", errr.cause.name);
			}
			log.error("code: ", errr.code);
			log.error("body: ", errr.body);
			log.error("cause: ", errr.cause);
			log.error("data: ", errr.data);
			log.error(err);
			if (causeName === "UnexpectedHttpStatus") {
				await this.puppet.sendStatusMessage(puppetId, "Error: " + err);
				await this.puppet.sendStatusMessage(puppetId, "Reconnecting in a minute... ");
				await this.stopClient(puppetId);
				setTimeout(async () => {
					await this.startClient(puppetId);
				}, MINUTE);
			} else {
				log.error("baaaad error");
				await this.puppet.sendStatusMessage(puppetId, "Super bad error, stopping puppet. Please restart the service to start it up again. This is for debugging purposes and will not be needed in the future");
				await this.stopClient(puppetId);
			}
		});
		try {
			await client.connect();
			await this.puppet.setUserId(puppetId, client.username);
			p.data.state = client.getState;
			await this.puppet.setPuppetData(puppetId, p.data);
			await this.puppet.sendStatusMessage(puppetId, "connected");
		} catch (err) {
			log.error("Failed to connect", err);
			await this.puppet.sendStatusMessage(puppetId, "Failed to connect, reconnecting in a minute... " + err);
			setTimeout(async () => {
				await this.startClient(puppetId);
			}, MINUTE);
		}
	}

	public async newPuppet(puppetId: number, data: any) {
		if (this.puppets[puppetId]) {
			await this.deletePuppet(puppetId);
		}
		const client = new Client(data.username, data.password);
		const TWO_MIN = 120000;
		this.puppets[puppetId] = {
			client,
			data,
			deletedMessages: new ExpireSet(TWO_MIN),
		};
		await this.startClient(puppetId);
	}

	public async deletePuppet(puppetId: number) {
		const p = this.puppets[puppetId];
		if (!p) {
			return;
		}
		await p.client.disconnect();
		delete this.puppets[puppetId];
	}

	public async createUser(remoteUser: IRemoteUser): Promise<IRemoteUser | null> {
		const p = this.puppets[remoteUser.puppetId];
		if (!p) {
			return null;
		}
		log.info(`Received create request for user update puppetId=${remoteUser.puppetId} userId=${remoteUser.userId}`);
		const contact = await p.client.getContact(remoteUser.userId);
		if (!contact) {
			return null;
		}
		return this.getUserParams(remoteUser.puppetId, contact);
	}

	public async createRoom(room: IRemoteRoom): Promise<IRemoteRoom | null> {
		const p = this.puppets[room.puppetId];
		if (!p) {
			return null;
		}
		log.info(`Received create request for channel update puppetId=${room.puppetId} roomId=${room.roomId}`);
		const conversation = await p.client.getConversation(room);
		if (!conversation) {
			return null;
		}
		return this.getRoomParams(room.puppetId, conversation);
	}

	public async getDmRoom(remoteUser: IRemoteUser): Promise<string | null> {
		const p = this.puppets[remoteUser.puppetId];
		if (!p) {
			return null;
		}
		const contact = await p.client.getContact(remoteUser.userId);
		if (!contact) {
			return null;
		}
		return `dm-${remoteUser.puppetId}-${contact.mri}`;
	}

	public async listUsers(puppetId: number): Promise<IRetList[]> {
		const p = this.puppets[puppetId];
		if (!p) {
			return [];
		}
		const reply: IRetList[] = [];
		for (const [, contact] of p.client.contacts) {
			if (!contact) {
				continue;
			}
			reply.push({
				id: contact.mri,
				name: contact.displayName,
			});
		}
		return reply;
	}

	public async listRooms(puppetId: number): Promise<IRetList[]> {
		const p = this.puppets[puppetId];
		if (!p) {
			return [];
		}
		const reply: IRetList[] = [];
		for (const [, conversation] of p.client.conversations) {
			if (!conversation || conversation.id.startsWith("8:")) {
				continue;
			}
			reply.push({
				id: conversation.id,
				name: (conversation.threadProperties && conversation.threadProperties.topic) || "",
			});
		}
		return reply;
	}

	public async getUserIdsInRoom(room: IRemoteRoom): Promise<Set<string> | null> {
		const p = this.puppets[room.puppetId];
		if (!p) {
			return null;
		}
		const conversation = await p.client.getConversation(room);
		if (!conversation) {
			return null;
		}
		const users = new Set<string>();
		if (conversation.members) {
			for (const member of conversation.members) {
				users.add(member);
			}
		}
		return users;
	}

	public async handleMatrixMessage(room: IRemoteRoom, data: IMessageEvent) {
		const p = this.puppets[room.puppetId];
		if (!p) {
			return;
		}
		log.info("Received message from matrix");
		const conversation = await p.client.getConversation(room);
		if (!conversation) {
			log.warn(`Room ${room.roomId} not found!`);
			return;
		}
		let msg: string;
		if (data.formattedBody) {
			msg = this.matrixMessageParser.parse(data.formattedBody);
		} else {
			msg = escapeHtml(data.body);
		}
		const dedupeKey = `${room.puppetId};${room.roomId}`;
		this.messageDeduplicator.lock(dedupeKey, p.client.username, msg);
		const ret = await p.client.sendMessage(conversation.id, msg);
		const eventId = ret && ret.MessageId;
		this.messageDeduplicator.unlock(dedupeKey, p.client.username, eventId);
		if (eventId) {
			await this.puppet.eventSync.insert(room.puppetId, data.eventId!, eventId);
		}
	}

	public async handleMatrixEdit(room: IRemoteRoom, eventId: string, data: IMessageEvent) {
		const p = this.puppets[room.puppetId];
		if (!p) {
			return;
		}
		log.info("Received edit from matrix");
		const conversation = await p.client.getConversation(room);
		if (!conversation) {
			log.warn(`Room ${room.roomId} not found!`);
			return;
		}
		let msg: string;
		if (data.formattedBody) {
			msg = this.matrixMessageParser.parse(data.formattedBody);
		} else {
			msg = escapeHtml(data.body);
		}
		const dedupeKey = `${room.puppetId};${room.roomId}`;
		this.messageDeduplicator.lock(dedupeKey, p.client.username, msg);
		await p.client.sendEdit(conversation.id, eventId, msg);
		const newEventId = "";
		this.messageDeduplicator.unlock(dedupeKey, p.client.username, newEventId);
		if (newEventId) {
			await this.puppet.eventSync.insert(room.puppetId, data.eventId!, newEventId);
		}
	}

	public async handleMatrixRedact(room: IRemoteRoom, eventId: string) {
		const p = this.puppets[room.puppetId];
		if (!p) {
			return;
		}
		log.info("Received edit from matrix");
		const conversation = await p.client.getConversation(room);
		if (!conversation) {
			log.warn(`Room ${room.roomId} not found!`);
			return;
		}
		p.deletedMessages.add(eventId);
		await p.client.sendDelete(conversation.id, eventId);
	}

	public async handleMatrixImage(room: IRemoteRoom, data: IFileEvent) {
		await this.handleMatrixFile(room, data, "sendImage");
	}

	public async handleMatrixAudio(room: IRemoteRoom, data: IFileEvent) {
		await this.handleMatrixFile(room, data, "sendAudio");
	}

	public async handleMatrixFile(room: IRemoteRoom, data: IFileEvent, method?: string) {
		if (!method) {
			method = "sendDocument";
		}
		const p = this.puppets[room.puppetId];
		if (!p) {
			return;
		}
		log.info("Received file from matrix");
		const conversation = await p.client.getConversation(room);
		if (!conversation) {
			log.warn(`Room ${room.roomId} not found!`);
			return;
		}
		const buffer = await Util.DownloadFile(data.url);
		const opts: SkypeNewMediaMessage = {
			file: buffer,
			name: data.filename,
		};
		if (data.info) {
			if (data.info.w) {
				opts.width = data.info.w;
			}
			if (data.info.h) {
				opts.height = data.info.h;
			}
		}
		const dedupeKey = `${room.puppetId};${room.roomId}`;
		this.messageDeduplicator.lock(dedupeKey, p.client.username, `file:${data.filename}`);
		const ret = await p.client[method](conversation.id, opts);
		const eventId = ret && ret.MessageId;
		this.messageDeduplicator.unlock(dedupeKey, p.client.username, eventId);
		if (eventId) {
			await this.puppet.eventSync.insert(room.puppetId, data.eventId!, eventId);
		}
	}

	private async handleSkypeText(
		puppetId: number,
		resource: skypeHttp.resources.TextResource | skypeHttp.resources.RichTextResource,
	) {
		const p = this.puppets[puppetId];
		if (!p) {
			return;
		}
		const rich = resource.native.messagetype.startsWith("RichText");
		log.info("Got new skype message");
		log.silly(resource);
		const params = await this.getSendParams(puppetId, resource);
		if (!params) {
			log.warn("Couldn't generate params");
			return;
		}
		let msg = resource.content;
		let emote = false;
		if (resource.native.skypeemoteoffset) {
			emote = true;
			msg = msg.substr(Number(resource.native.skypeemoteoffset));
		}
		const dedupeKey = `${puppetId};${params.room.roomId}`;
		if (rich && msg.trim().startsWith("<URIObject") && msg.trim().endsWith("</URIObject>")) {
			// okay, we might have a sticker or something...
			const $ = cheerio.load(msg);
			const obj = $("URIObject");
			let uri = obj.attr("uri");
			const filename = $(obj.find("OriginalName")).attr("v");
			if (uri) {
				if (await this.messageDeduplicator.dedupe(dedupeKey, params.user.userId, params.eventId, `file:${filename}`)) {
					log.silly("file message dedupe");
					return;
				}
				uri += "/views/thumblarge";
				uri = uri.replace("static.asm.skype.com", "static-asm.secure.skypeassets.com");
				const buffer = await p.client.downloadFile(uri);
				await this.puppet.sendFileDetect(params, buffer, filename);
				return;
			}
		}
		if (await this.messageDeduplicator.dedupe(dedupeKey, params.user.userId, params.eventId, msg)) {
			log.silly("normal message dedupe");
			return;
		}
		let sendMsg: IMessageEvent;
		if (rich) {
			sendMsg = this.skypeMessageParser.parse(msg);
		} else {
			sendMsg = {
				body: msg,
			};
		}
		if (emote) {
			sendMsg.emote = true;
		}
		await this.puppet.sendMessage(params, sendMsg);
	}

	private async handleSkypeEdit(
		puppetId: number,
		resource: skypeHttp.resources.TextResource | skypeHttp.resources.RichTextResource,
	) {
		const p = this.puppets[puppetId];
		if (!p) {
			return;
		}
		const rich = resource.native.messagetype.startsWith("RichText");
		log.info("Got new skype edit");
		log.silly(resource);
		const params = await this.getSendParams(puppetId, resource);
		if (!params) {
			log.warn("Couldn't generate params");
			return;
		}
		let msg = resource.content;
		let emote = false;
		if (resource.native.skypeemoteoffset) {
			emote = true;
			msg = msg.substr(Number(resource.native.skypeemoteoffset));
		}
		const dedupeKey = `${puppetId};${params.room.roomId}`;
		if (await this.messageDeduplicator.dedupe(dedupeKey, params.user.userId, params.eventId, msg)) {
			log.silly("normal message dedupe");
			return;
		}
		let sendMsg: IMessageEvent;
		if (rich) {
			sendMsg = this.skypeMessageParser.parse(msg);
		} else {
			sendMsg = {
				body: msg,
			};
		}
		if (emote) {
			sendMsg.emote = true;
		}
		if (resource.content) {
			await this.puppet.sendEdit(params, resource.id, sendMsg);
		} else if (p.deletedMessages.has(resource.id)) {
			log.silly("normal message redact dedupe");
			return;
		} else {
			await this.puppet.sendRedact(params, resource.id);
		}
	}

	private async handleSkypeFile(puppetId: number, resource: skypeHttp.resources.FileResource) {
		const p = this.puppets[puppetId];
		if (!p) {
			return;
		}
		log.info("Got new skype file");
		log.silly(resource);
		const params = await this.getSendParams(puppetId, resource);
		if (!params) {
			log.warn("Couldn't generate params");
			return;
		}
		const filename = resource.original_file_name;
		const dedupeKey = `${puppetId};${params.room.roomId}`;
		if (await this.messageDeduplicator.dedupe(dedupeKey, params.user.userId, params.eventId, `file:${filename}`)) {
			log.silly("file message dedupe");
			return;
		}
		const buffer = await p.client.downloadFile(resource.uri);
		await this.puppet.sendFileDetect(params, buffer, filename);
	}

	private async handleSkypeTyping(puppetId: number, resource: skypeHttp.resources.Resource, typing: boolean) {
		const p = this.puppets[puppetId];
		if (!p) {
			return;
		}
		log.info("Got new skype typing event");
		log.silly(resource);
		const params = await this.getSendParams(puppetId, resource);
		if (!params) {
			log.warn("Couldn't generate params");
			return;
		}
		await this.puppet.setUserTyping(params, typing);
	}

	private async handleSkypePresence(puppetId: number, resource: skypeHttp.resources.Resource) {
		const p = this.puppets[puppetId];
		if (!p) {
			return;
		}
		log.info("Got new skype presence event");
		log.silly(resource);
		const content = JSON.parse(resource.native.content);
		const contact = await p.client.getContact(content.user);
		const conversation = await p.client.getConversation({
			puppetId,
			roomId: resource.conversation,
		});
		if (!contact || !conversation) {
			log.warn("Couldn't generate params");
			return;
		}
		const params: IReceiveParams = {
			user: this.getUserParams(puppetId, contact),
			room: this.getRoomParams(puppetId, conversation),
		};
		const [id, _, clientId] = content.consumptionhorizon.split(";");
		params.eventId = id;
		await this.puppet.sendReadReceipt(params);
		params.eventId = clientId;
		await this.puppet.sendReadReceipt(params);
	}
}
