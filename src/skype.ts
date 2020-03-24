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
	ExpireSet,
} from "mx-puppet-bridge";
import { Client } from "./client";
import * as skypeHttp from "skype-http";
import { Contact as SkypeContact } from "skype-http/dist/lib/types/contact";
import { NewMediaMessage as SkypeNewMediaMessage } from "skype-http/dist/lib/interfaces/api/api";
import * as decodeHtml from "decode-html";
import * as escapeHtml from "escape-html";

const log = new Log("SkypePuppet:skype");

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
	constructor(
		private puppet: PuppetBridge,
	) {
		this.messageDeduplicator = new MessageDeduplicator();
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
		let roomId = conversation.id;
		const isDirect = roomType === 8;
		if (isDirect) {
			return {
				puppetId,
				roomId: `dm-${puppetId}-${roomId}`,
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
			roomId,
			name,
			avatarUrl,
		};
	}

	public async getSendParams(puppetId: number, resource: skypeHttp.resources.Resource): Promise<IReceiveParams | null> {
		const roomType = Number(resource.conversation.split(":")[0]);
		let roomId = resource.conversation;
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
			user: await this.getUserParams(puppetId, contact),
			room: await this.getRoomParams(puppetId, conversation),
			eventId: (resource as any).clientId || resource.native.clientmessageid || resource.id, // tslint:disable-line no-any
		};
	}

	public async startClient(puppetId: number) {
		const p = this.puppets[puppetId];
		if (!p) {
			return;
		}
		const client = p.client;
		client.on("text", async (resource: skypeHttp.resources.TextResource) => {
			try {
				await this.handleSkypeText(puppetId, resource, false);
			} catch (err) {
				log.error("Error while handling text event", err);
			}
		});
		client.on("richText", async (resource: skypeHttp.resources.RichTextResource) => {
			try {
				await this.handleSkypeText(puppetId, resource, true);
			} catch (err) {
				log.error("Error while handling richText event", err);
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
				
			} catch (err) {
				log.error("Error while handling typing event", err);
			}
		});
		await client.connect();
		await this.puppet.setUserId(puppetId, client.username);
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
			msg = data.formattedBody;
		} else {
			msg = escapeHtml(data.body);
		}
		const dedupeKey = `${room.puppetId};${room.roomId}`;
		this.messageDeduplicator.lock(dedupeKey, p.client.username, msg);
		const ret = await p.client.sendMessage(conversation.id, msg);
		const dedupeId = ret && ret.clientMessageId;
		const eventId = ret && ret.MessageId;
		this.messageDeduplicator.unlock(dedupeKey, p.client.username, dedupeId);
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
			msg = data.formattedBody;
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
		const dedupeId = ret && ret.clientMessageId;
		const eventId = ret && ret.MessageId;
		this.messageDeduplicator.unlock(dedupeKey, p.client.username, dedupeId);
		if (eventId) {
			await this.puppet.eventSync.insert(room.puppetId, data.eventId!, eventId);
		}
	}

	private async handleSkypeText(
		puppetId: number,
		resource: skypeHttp.resources.TextResource | skypeHttp.resources.RichTextResource,
		rich: boolean,
	) {
		const p = this.puppets[puppetId];
		if (!p) {
			return;
		}
		log.info("Got new skype message");
		log.silly(resource);
		const params = await this.getSendParams(puppetId, resource);
		if (!params) {
			log.warn("Couldn't generate params");
			return;
		}
		let msg = resource.content;
		let emote = false;
		if (resource.native && resource.native.skypeemoteoffset) {
			emote = true;
			msg = msg.substr(Number(resource.native.skypeemoteoffset));
		}
		const dedupeKey = `${puppetId};${params.room.roomId}`;
		if (await this.messageDeduplicator.dedupe(dedupeKey, params.user.userId, params.eventId, msg)) {
			log.silly("normal message dedupe");
			return;
		}
		if (!rich) {
			await this.puppet.sendMessage(params, {
				body: msg,
				emote,
			});
		} else if (resource.native && resource.native.skypeeditedid) {
			if (resource.content) {
				await this.puppet.sendEdit(params, resource.native.skypeeditedid, {
					body: msg,
					formattedBody: msg,
					emote,
				});
			} else if (p.deletedMessages.has(resource.native.skypeeditedid)) {
				log.silly("normal message redact dedupe");
				return;
			} else {
				await this.puppet.sendRedact(params, resource.native.skypeeditedid);
			}
		} else {
			await this.puppet.sendMessage(params, {
				body: msg,
				formattedBody: msg,
				emote,
			});
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
		if (resource.native && resource.native.skypeeditedid && !resource.uri) {
			if (p.deletedMessages.has(resource.native.skypeeditedid)) {
				log.silly("file message redact dedupe");
				return;
			}
			await this.puppet.sendRedact(params, resource.native.skypeeditedid);
			return;
		}
		const buffer = await p.client.downloadFile(resource.uri);
		await this.puppet.sendFileDetect(params, buffer, filename);
	}
}
