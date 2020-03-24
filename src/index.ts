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
	PuppetBridge,
	IPuppetBridgeRegOpts,
	Log,
	IRetData,
	Util,
	IProtocolInformation,
} from "mx-puppet-bridge";
import * as commandLineArgs from "command-line-args";
import * as commandLineUsage from "command-line-usage";
import * as fs from "fs";
import * as yaml from "js-yaml";
import { Skype } from "./skype";
import { Client } from "./client";

const log = new Log("SkypePuppet:index");

const commandOptions = [
	{ name: "register", alias: "r", type: Boolean },
	{ name: "registration-file", alias: "f", type: String },
	{ name: "config", alias: "c", type: String },
	{ name: "help", alias: "h", type: Boolean },
];
const options = Object.assign({
	"register": false,
	"registration-file": "skype-registration.yaml",
	"config": "config.yaml",
	"help": false,
}, commandLineArgs(commandOptions));

if (options.help) {
	// tslint:disable-next-line:no-console
	console.log(commandLineUsage([
		{
			header: "Matrix Skype Puppet Bridge",
			content: "A matrix puppet bridge for skype",
		},
		{
			header: "Options",
			optionList: commandOptions,
		},
	]));
	process.exit(0);
}

const protocol: IProtocolInformation = {
	features: {
//		file: true, // no need for the others as we auto-detect types anyways
//		presence: true, // we want to be able to send presence
		globalNamespace: true,
	},
	id: "skype",
	displayname: "Skype",
	externalUrl: "https://skype.com/",
};

const puppet = new PuppetBridge(options["registration-file"], options.config, protocol);

if (options.register) {
	// okay, all we have to do is generate a registration file
	puppet.readConfig(false);
	try {
		puppet.generateRegistration({
			prefix: "_skypepuppet_",
			id: "skype-puppet",
			url: `http://${puppet.Config.bridge.bindAddress}:${puppet.Config.bridge.port}`,
		});
	} catch (err) {
		// tslint:disable-next-line:no-console
		console.log("Couldn't generate registration file:", err);
	}
	process.exit(0);
}

async function run() {
	await puppet.init();
	const skype = new Skype(puppet);
	puppet.on("puppetNew", skype.newPuppet.bind(skype));
	puppet.on("puppetDelete", skype.deletePuppet.bind(skype));
	puppet.on("message", skype.handleMatrixMessage.bind(skype));
	puppet.setCreateUserHook(skype.createUser.bind(skype));
	puppet.setCreateRoomHook(skype.createRoom.bind(skype));
	puppet.setGetUserIdsInRoomHook(skype.getUserIdsInRoom.bind(skype));
	puppet.setGetDataFromStrHook(async (str: string): Promise<IRetData> => {
		const retData = {
			success: false,
		} as IRetData;
		const [username, password] = str.split(" ");
		try {
			const client = new Client(username, password);
			await client.connect();
			await client.disconnect();
		} catch (err) {
			retData.error = "Username or password wrong";
			return retData;
		}
		retData.success = true;
		const data: any = {
			username,
			password,
		};
		retData.data = data;
		return retData;
	});
	puppet.setBotHeaderMsgHook((): string => {
		return "Skype Puppet Bridge";
	});
	await puppet.start();
}

// tslint:disable-next-line:no-floating-promises
run();
