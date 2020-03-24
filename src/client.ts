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

import { Log, ExpireSet, IRemoteRoom, Util } from "mx-puppet-bridge";
import { EventEmitter } from "events";
import * as skypeHttp from "skype-http";
import { Contact as SkypeContact } from "skype-http/dist/lib/types/contact";
import { NewMediaMessage as SkypeNewMediaMessage } from "skype-http/dist/lib/interfaces/api/api";
import { Context as SkypeContext } from "skype-http/dist/lib/interfaces/api/context";

const log = new Log("SkypePuppet:client");

const ID_TIMEOUT = 60000;

export class Client extends EventEmitter {
	public contacts: Map<string, SkypeContact | null> = new Map();
	public conversations: Map<string, skypeHttp.Conversation | null> = new Map();
	private api: skypeHttp.Api;
	private handledIds: ExpireSet<string>;
	constructor(
		private loginUsername: string,
		private password: string,
		private state?: SkypeContext.Json,
	) {
		super();
		this.handledIds = new ExpireSet(ID_TIMEOUT);
	}

	public get username(): string {
		return "8:" + this.api.context.username;
	}

	public get getState(): SkypeContext.Json {
		return this.api.getState();
	}

	public async connect() {
		if (this.state) {
			try {
				this.api = await skypeHttp.connect({ state: this.state, verbose: true });
			} catch (err) {
				this.api = await skypeHttp.connect({
					credentials: {
						username: this.loginUsername,
						password: this.password,
					},
					verbose: true,
				});
			}
		} else {
			this.api = await skypeHttp.connect({
				credentials: {
					username: this.loginUsername,
					password: this.password,
				},
				verbose: true,
			});
		}

		this.api.on("event", (evt: skypeHttp.events.EventMessage) => {
			if (!evt || !evt.resource) {
				return;
			}
			const resource = evt.resource;
			if (this.handledIds.has(resource.id)) {
				return;
			}
			this.handledIds.add(resource.id);
			log.debug(`Got new event of type ${resource.type}`);
			const [type, subtype] = resource.type.split("/");
			switch (type) {
				case "Text":
					this.emit("text", resource);
					break;
				case "RichText":
					if (subtype === "Location") {
						this.emit("location", resource);
					} else if (subtype) {
						this.emit("file", resource);
					} else {
						this.emit("richText", resource);
					}
					break;
				case "Control":
					if (subtype === "Typing" || subtype === "ClearTyping") {
						this.emit("typing", resource, subtype === "Typing");
					}
					break;
				case "ThreadActivity":
					if (subtype === "MemberConsumptionHorizonUpdate") {
						this.emit("presence", resource);
					}
					break;
			}
		});

		this.api.on("error", (err: Error) => {
			log.error("An error occured", err);
			this.emit("error", err);
		});

		await this.api.listen();
		await this.api.setStatus("Online");

		try {
			const contacts = await this.api.getContacts();
			for (const contact of contacts) {
				this.contacts.set(contact.mri, contact);
			}
			const conversations = await this.api.getConversations();
			for (const conversation of conversations) {
				this.conversations.set(conversation.id, conversation);
			}
		} catch (err) {
			log.error(err);
		}
	}

	public async disconnect() {
		await this.api.stopListening();
	}

	public async getContact(id: string): Promise<SkypeContact | null> {
		const hasStart = Boolean(id.match(/^\d+:/));
		const fullId = hasStart ? id : `8:${id}`;
		if (this.contacts.has(fullId)) {
			return this.contacts.get(fullId) || null;
		}
		if (hasStart) {
			id = id.substr(id.indexOf(":") + 1);
		}
		try {
			const rawContact = await this.api.getContact(id);
			const contact: SkypeContact = {
				personId: rawContact.id.raw,
				workloads: null,
				mri: rawContact.id.raw,
				blocked: false,
				authorized: true,
				creationTime: new Date(),
				displayName: (rawContact.name && rawContact.name.displayName) || rawContact.id.id,
				displayNameSource: "profile" as any, // tslint:disable-line no-any
				profile: {
					avatarUrl: rawContact.avatarUrl || undefined,
					name: {
						first: (rawContact.name && rawContact.name).first || undefined,
						surname: (rawContact.name && rawContact.name).surname || undefined,
						nickname: (rawContact.name && rawContact.name).nickname || undefined,
					},
				},
			};
			this.contacts.set(contact.mri, contact || null);
			return contact || null;
		} catch (err) {
			// contact not found
			this.contacts.set(fullId, null);
			return null;
		}
	}

	public async getConversation(room: IRemoteRoom): Promise<skypeHttp.Conversation | null> {
		let id = room.roomId;
		const match = id.match(/^dm-\d+-/);
		if (match) {
			const [_, puppetId] = id.split("-");
			if (Number(puppetId) !== room.puppetId) {
				return null;
			}
			id = id.substr(match[0].length);
		}
		if (this.conversations.has(id)) {
			return this.conversations.get(id) || null;
		}
		try {
			const conversation = await this.api.getConversation(id);
			this.conversations.set(conversation.id, conversation || null);
			return conversation || null;
		} catch (err) {
			// conversation not found
			this.conversations.set(id, null);
			return null;
		}
	}

	public async downloadFile(url: string): Promise<Buffer> {
		if (!url.includes("/views/imgpsh_fullsize_anim")) {
			url = url + "/views/imgpsh_fullsize_anim";
		}
		return await Util.DownloadFile(url, {
			cookies: this.api.context.cookies,
			headers: { Authorization: "skype_token " + this.api.context.skypeToken.value },
		});
	}

	public async sendMessage(conversationId: string, msg: string): Promise<skypeHttp.Api.SendMessageResult> {
		return await this.api.sendMessage({
			textContent: msg,
		}, conversationId);
	}

	public async sendEdit(conversationId: string, messageId: string, msg: string) {
		return await this.api.sendEdit({
			textContent: msg,
		}, conversationId, messageId);
	}

	public async sendDelete(conversationId: string, messageId: string) {
		return await this.api.sendDelete(conversationId, messageId);
	}

	public async sendAudio(
		conversationId: string,
		opts: SkypeNewMediaMessage,
	): Promise<skypeHttp.Api.SendMessageResult> {
		return await this.api.sendAudio(opts, conversationId);
	}

	public async sendDocument(
		conversationId: string,
		opts: SkypeNewMediaMessage,
	): Promise<skypeHttp.Api.SendMessageResult> {
		return await this.api.sendDocument(opts, conversationId);
	}

	public async sendImage(
		conversationId: string,
		opts: SkypeNewMediaMessage,
	): Promise<skypeHttp.Api.SendMessageResult> {
		return await this.api.sendImage(opts, conversationId);
	}
}
