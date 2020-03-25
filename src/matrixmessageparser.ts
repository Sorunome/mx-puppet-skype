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
import * as escapeHtml from "escape-html";

export class MatrixMessageParser {
	public parse(msg: string): string {
		const nodes = Parser.parse(`<wrap>${msg}</wrap>`, {
			lowerCaseTagName: true,
			pre: true,
		});
		return this.walkNode(nodes);
	}

	private walkChildNodes(node: Parser.Node): string {
		return node.childNodes.map((n) => this.walkNode(n)).join("");
	}

	private escape(s: string): string {
		return s;
	}

	private walkNode(node: Parser.Node): string {
		if (node.nodeType === Parser.NodeType.TEXT_NODE) {
			return this.escape((node as Parser.TextNode).text);
		} else if (node.nodeType === Parser.NodeType.ELEMENT_NODE) {
			const nodeHtml = node as Parser.HTMLElement;
			switch (nodeHtml.tagName) {
				case "em":
				case "i":
					return `<i raw_pre="_" raw_post="_">${this.walkChildNodes(nodeHtml)}</i>`;
				case "strong":
				case "b":
					return `<b raw_pre="*" raw_post="*">${this.walkChildNodes(nodeHtml)}</b>`;
				case "del":
					return `<s raw_pre="~" raw_post="~">${this.walkChildNodes(nodeHtml)}</s>`;
				case "code":
					return `<pre raw_pre="{code}" raw_post="{code}">${this.walkChildNodes(nodeHtml)}</pre>`;
				case "a": {
					const href = nodeHtml.attributes.href;
					const inner = this.walkChildNodes(nodeHtml);
					return `<a href="${escapeHtml(href)}">${inner}</a>`;
				}
				case "wrap":
					return this.walkChildNodes(nodeHtml);
				default:
					if (!nodeHtml.tagName) {
						return this.walkChildNodes(nodeHtml);
					}
					return `<${nodeHtml.tagName}>${this.walkChildNodes(nodeHtml)}</${nodeHtml.tagName}>`;
			}
		}
		return "";
	}
}
