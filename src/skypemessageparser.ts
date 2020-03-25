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

export class SkypeMessageParser {
	public parse(msg: string): IMessageEvent {
		const nodes = Parser.parse(`<wrap>${msg}</wrap>`, {
			lowerCaseTagName: true,
			pre: true,
		});
		return this.walkNode(nodes);
	}

	private walkChildNodes(node: Parser.Node): IMessageEvent {
		return node.childNodes.map((node) => this.walkNode(node)).reduce((acc, curr) => {
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

	private walkNode(node: Parser.Node): IMessageEvent {
		if (node.nodeType === Parser.NodeType.TEXT_NODE) {
			return this.escape((node as Parser.TextNode).text);
		} else if (node.nodeType === Parser.NodeType.ELEMENT_NODE) {
			const nodeHtml = node as Parser.HTMLElement;
			switch (nodeHtml.tagName) {
				case "i": {
					const child = this.walkChildNodes(nodeHtml);
					return {
						body: `_${child.body}_`,
						formattedBody: `<em>${child.formattedBody}</em>`,
					};
				}
				case "b": {
					const child = this.walkChildNodes(nodeHtml);
					return {
						body: `*${child.body}*`,
						formattedBody: `<strong>${child.formattedBody}</strong>`,
					};
				}
				case "s": {
					const child = this.walkChildNodes(nodeHtml);
					return {
						body: `~${child.body}~`,
						formattedBody: `<del>${child.formattedBody}</del>`,
					};
				}
				case "pre": {
					const child = this.walkChildNodes(nodeHtml);
					return {
						body: `{code}${child.body}{code}`,
						formattedBody: `<code>${child.formattedBody}</code>`,
					};
				}
				case "a": {
					const href = nodeHtml.attributes.href;
					const child = this.walkChildNodes(nodeHtml);
					return {
						body: child.body === href ? href : `[${child.body}](${href})`,
						formattedBody: `<a href="${escapeHtml(href)}">${child.formattedBody}</a>`,
					};
				}
				case "e_m":
					return {
						body: "",
						formattedBody: "",
					};
				default:
					return this.walkChildNodes(nodeHtml);
			}
		}
		return {
			body: "",
			formattedBody: "",
		};
	}
}
