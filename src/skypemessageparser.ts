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

import * as Parser from "node-html-parser";
import * as decodeHtml from "decode-html";
import * as escapeHtml from "escape-html";
import { IMessageEvent } from "mx-puppet-bridge";
import * as emoji from "node-emoji";

interface ISkypeMessageParserOpts {
	noQuotes?: boolean;
}

export class SkypeMessageParser {
	public parse(msg: string, opts: ISkypeMessageParserOpts = {}): IMessageEvent {
		opts = Object.assign({
			noQuotes: false,
		}, opts);
		const nodes = Parser.parse(`<wrap>${msg}</wrap>`, {
			lowerCaseTagName: true,
			pre: true,
		});
		return this.walkNode(nodes, opts);
	}

	private walkChildNodes(node: Parser.Node, opts: ISkypeMessageParserOpts): IMessageEvent {
		if (!node.childNodes.length) {
			return {
				body: "",
				formattedBody: "",
			};
		}
		return node.childNodes.map((n) => this.walkNode(n, opts)).reduce((acc, curr) => {
			return {
				body: acc.body + curr.body,
				formattedBody: acc.formattedBody! + curr.formattedBody!,
			};
		});
	}

	private escape(s: string): IMessageEvent {
		return {
			body: decodeHtml(s),
			formattedBody: s.replace("\n", "<br>"),
		};
	}

	private walkNode(node: Parser.Node, opts: ISkypeMessageParserOpts): IMessageEvent {
		if (node.nodeType === Parser.NodeType.TEXT_NODE) {
			return this.escape((node as Parser.TextNode).text);
		} else if (node.nodeType === Parser.NodeType.ELEMENT_NODE) {
			const nodeHtml = node as Parser.HTMLElement;
			switch (nodeHtml.tagName) {
				case "i": {
					const child = this.walkChildNodes(nodeHtml, opts);
					return {
						body: `_${child.body}_`,
						formattedBody: `<em>${child.formattedBody}</em>`,
					};
				}
				case "b": {
					const child = this.walkChildNodes(nodeHtml, opts);
					return {
						body: `*${child.body}*`,
						formattedBody: `<strong>${child.formattedBody}</strong>`,
					};
				}
				case "s": {
					const child = this.walkChildNodes(nodeHtml, opts);
					return {
						body: `~${child.body}~`,
						formattedBody: `<del>${child.formattedBody}</del>`,
					};
				}
				case "pre": {
					const child = this.walkChildNodes(nodeHtml, opts);
					return {
						body: `{code}${child.body}{code}`,
						formattedBody: `<code>${child.formattedBody}</code>`,
					};
				}
				case "a": {
					const href = nodeHtml.attributes.href;
					const child = this.walkChildNodes(nodeHtml, opts);
					return {
						body: child.body === href ? href : `[${child.body}](${href})`,
						formattedBody: `<a href="${escapeHtml(href)}">${child.formattedBody}</a>`,
					};
				}
				case "quote": {
					if (opts.noQuotes) {
						return {
							body: "",
							formattedBody: "",
						};
					}
					const child = this.walkChildNodes(nodeHtml, opts);
					return {
						body: `> ${child.body}\n`,
						formattedBody: `<blockquote>${child.formattedBody}<br> - ${nodeHtml.attributes.authorname}</blockquote>`,
					};
				}
				case "ss": {
					// skype emoji
					const type = nodeHtml.attributes.type;
					let emojiType = {
						smile: "slightly_smiling_face",
						sad: "slightly_frowning_face",
						laugh: "grin",
						cool: "sunglasses",
						hearteyes: "heart_eyes",
						stareyes: "star-struck",
						like: "thumbsup",
						cwl: "rolling_on_the_floor_laughing",
						xd: "laughing",
						happyface: "smiley",
						happyeyes: "smile",
						// hysterical: "", TODO: find
						sweatgrinning: "sweat_smile",
						// smileeyes: "", TODO: find
						blankface: "no_mouth",
						surprised: "astonished",
						upsidedownface: "upside_down_face",
						loudlycrying: "sob",
						shivering: "🥶",
						speechless: "😐️",
						tongueout: "stuck_out_tongue",
						winktongueout: "stuck_out_tongue_winking_eye",
						inlove: "🥰",
						// wonder: "", TODO: find
						// dull: "", TODO: find
						yawn: "🥱",
						puke: "face_vomiting",
						// doh: "", TODO: find
						angryface: "angry",
						angry: "rage",
						// wasntme: "", TODO: find
						// worry: "", TODO: find
						screamingfear: "scream",
						// veryconfused: "", TODO: find
						// mmm: "", TODO: find
						nerdy: "nerd_face",
						loveearth: "🌍️",
						// rainbowsmile: "", TODO: find
						// lipssealed: "", TODO: find
						devil: "smiling_imp",
						// envy: "", TODO: find
						// makeup: "", TODO: find
						think: "thinking_face",
						rofl: "rolling_on_the_floor_laughing",
					}[type];
					const haveEmojiType = Boolean(emojiType);
					if (!emojiType) {
						emojiType = type;
					}
					let e = emoji.get(emojiType);
					if (e === `:${emojiType}:`) {
						e = emoji.get(emojiType + "_face");
					}
					if (!e.startsWith(":")) {
						return {
							body: e,
							formattedBody: e,
						};
					}
					if (haveEmojiType) {
						return {
							body: emojiType,
							formattedBody: emojiType,
						};
					}
					return {
						body: `(${type})`,
						formattedBody: `(${escapeHtml(type)})`,
					};
				}
				case "e_m": // empty edit tag
				case "legacyquote": // empty legacy quote tag
					return {
						body: "",
						formattedBody: "",
					};
				default:
					return this.walkChildNodes(nodeHtml, opts);
			}
		}
		return {
			body: "",
			formattedBody: "",
		};
	}
}
