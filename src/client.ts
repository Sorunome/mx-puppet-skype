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

import { Log, IRemoteRoom, Util } from "mx-puppet-bridge";
import { EventEmitter } from "events";
import * as skypeHttp from "skype-http";
import { Contact as SkypeContact } from "skype-http/dist/lib/types/contact";
import { NewMediaMessage as SkypeNewMediaMessage } from "skype-http/dist/lib/interfaces/api/api";
import { Context as SkypeContext } from "skype-http/dist/lib/interfaces/api/context";
import ExpireSet from "expire-set";

const log = new Log("SkypePuppet:client");

// tslint:disable no-magic-numbers
const ID_TIMEOUT = 60000;
const CONTACTS_DELTA_INTERVAL = 5 * 60 * 1000;
// tslint:enable no-magic-numbers

export class Client extends EventEmitter {
	public contacts: Map<string, SkypeContact | null> = new Map();
	public conversations: Map<string, skypeHttp.Conversation | null> = new Map();
	private api: skypeHttp.Api;
	private handledIds: ExpireSet<string>;
	private contactsInterval: NodeJS.Timeout | null = null;
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
		let connectedWithAuth = false;
		if (this.state) {
			try {
				this.api = await skypeHttp.connect({ state: this.state, verbose: true });
				connectedWithAuth = true;
			} catch (err) {
				this.api = await skypeHttp.connect({
					credentials: {
						username: this.loginUsername,
						password: this.password,
					},
					verbose: true,
				});
				connectedWithAuth = false;
			}
		} else {
			this.api = await skypeHttp.connect({
				credentials: {
					username: this.loginUsername,
					password: this.password,
				},
				verbose: true,
			});
			connectedWithAuth = false;
		}

		try {
			await this.startupApi();
		} catch (err) {
			if (!connectedWithAuth) {
				throw err;
			}
			this.api = await skypeHttp.connect({
				credentials: {
					username: this.loginUsername,
					password: this.password,
				},
				verbose: true,
			});
			connectedWithAuth = false;
			await this.startupApi();
		}

		const registerErrorHandler = () => {
			this.api.on("error", (err: Error) => {
				log.error("An error occured", err);
				this.emit("error", err);
			});
		};

		if (connectedWithAuth) {
			let resolved = false;
			return new Promise(async (resolve, reject) => {
				const TIMEOUT_SUCCESS = 5000;
				setTimeout(() => {
					if (resolved) {
						return;
					}
					resolved = true;
					registerErrorHandler();
					resolve();
				}, TIMEOUT_SUCCESS);
				this.api.once("error", async () => {
					if (resolved) {
						return;
					}
					resolved = true;
					// alright, re-try as normal user
					try {
						await this.api.stopListening();
						this.api = await skypeHttp.connect({
							credentials: {
								username: this.loginUsername,
								password: this.password,
							},
							verbose: true,
						});
						await this.startupApi();
						registerErrorHandler();
						resolve();
					} catch (err) {
						reject(err);
					}
				});
				await this.api.listen();
			}).then(async () => {
				await this.api.setStatus("Online");
			});
		} else {
			registerErrorHandler();
			await this.api.listen();
			await this.api.setStatus("Online");
		}
	}

	public async disconnect() {
		if (this.api) {
			await this.api.stopListening();
		}
		if (this.contactsInterval) {
			clearInterval(this.contactsInterval);
			this.contactsInterval = null;
		}
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
		if (!url.includes("/views/")) {
			url = url + "/views/imgpsh_fullsize_anim";
		}
		return await Util.DownloadFile(url, {
			cookies: this.api.context.cookies,
			headers: {
				Authorization: "skypetoken=" + this.api.context.skypeToken.value,
				RegistrationToken: this.api.context.registrationToken.raw,
			},
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

	private async startupApi() {
		this.api.on("event", (evt: skypeHttp.events.EventMessage) => {
			if (!evt || !evt.resource) {
				return;
			}
			const resource = evt.resource;
			log.debug(`Got new event of type ${resource.type}`);
			log.silly(evt);
			const [type, subtype] = resource.type.split("/");
			switch (type) {
				case "RichText":
					if (evt.resourceType === "NewMessage") {
						if (resource.native.skypeeditedid || this.handledIds.has(resource.id)) {
							break;
						}
						this.handledIds.add(resource.id);
						if (subtype === "Location") {
							this.emit("location", resource);
						} else if (subtype) {
							this.emit("file", resource);
						} else {
							this.emit("text", resource);
						}
					} else if (evt.resourceType === "MessageUpdate") {
						this.emit("edit", resource);
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

		const contacts = await this.api.getContacts();
		for (const contact of contacts) {
			this.contacts.set(contact.mri, contact);
		}
		const conversations = await this.api.getConversations();
		for (const conversation of conversations) {
			this.conversations.set(conversation.id, conversation);
		}

		if (this.contactsInterval) {
			clearInterval(this.contactsInterval);
			this.contactsInterval = null;
		}
		this.contactsInterval = setInterval(this.updateContacts.bind(this), CONTACTS_DELTA_INTERVAL);
	}

	private async updateContacts() {
		log.verbose("Getting contacts diff....");
		try {
			const contacts = await this.api.getContacts(true);
			const MANY_CONTACTS = 5;
			for (const contact of contacts) {
				const oldContact = this.contacts.get(contact.mri) || null;
				this.contacts.set(contact.mri, contact);
				this.emit("updateContact", oldContact, contact);
			}
		} catch (err) {
			log.error("Failed to get contacts diff", err);
			this.emit("error", err);
		}
	}
}
