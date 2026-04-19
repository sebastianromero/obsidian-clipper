import dayjs from "dayjs";
import { createMarkdownContent } from "defuddle/full";
import type { ExtractedContent } from "../types/types";
import browser from "./browser-polyfill";
import { debugLog } from "./debug";
import {
	getElementByXPath,
	wrapElementWithMark,
	wrapTextWithMark
} from "./dom-utils";
import type {
	AnyHighlightData,
	HighlightData,
	TextHighlightData
} from "./highlighter";
import { addSchemaOrgDataToVariables, buildVariables } from "./shared";
import { generalSettings } from "./storage-utils";
import { sanitizeFileName } from "./string-utils";

// Define ElementHighlightData type inline since it is not exported from highlighter.ts
interface ElementHighlightData extends HighlightData {
	type: "element";
}

function cleanAndFormatFigures(html: string): string {
	if (!html) return html;
	const parser = new DOMParser();
	const doc = parser.parseFromString(html, "text/html");

	// 1. Global cleanup: remove aria-hidden attributes and irrelevant interactive UI elements
	doc.querySelectorAll("[aria-hidden]").forEach((el) => {
		el.removeAttribute("aria-hidden");
	});
	doc
		.querySelectorAll("button, .btn, .figure-pop-btn, i.icon, .icon")
		.forEach((el) => {
			el.remove();
		});

	// 2. Process Figures to extract captions safely and enforce semantic markdown conversion
	doc.querySelectorAll("figure").forEach((figure) => {
		const img = figure.querySelector("img");
		const table = figure.querySelector("table");
		const figcaption = figure.querySelector("figcaption");
		const notes = figure.querySelector('.notes, [role="doc-footnote"]');

		let captionText =
			figcaption?.textContent?.trim().replace(/\s+/g, " ") || "";

		if (notes) {
			const notesText = notes.textContent?.trim().replace(/\s+/g, " ");
			if (notesText) {
				captionText = captionText ? `${captionText} (${notesText})` : notesText;
			}
		}

		if (img) {
			// Inject the evaluated caption directly into the alt attribute for images
			if (captionText) {
				img.setAttribute("alt", captionText);
			}
			figure.replaceWith(img);
		} else if (table) {
			// Create a semantic paragraph element for table captions to satisfy Pandoc requirements
			const fragment = doc.createDocumentFragment();
			if (captionText) {
				const captionPara = doc.createElement("p");
				captionPara.textContent = `Table: ${captionText}`;
				fragment.appendChild(captionPara);
			}
			fragment.appendChild(table);
			figure.replaceWith(fragment);
		}
	});

	return doc.body.innerHTML;
}

function protectComplexTables(html: string, protectedTables: string[]): string {
	if (!html) return html;
	const parser = new DOMParser();
	const doc = parser.parseFromString(html, "text/html");

	doc.querySelectorAll("table").forEach((table) => {
		const thead = table.querySelector("thead");

		// Determine if the table breaches standard markdown constraints (multi-row headers or merged cells)
		const hasMultipleHeaders = thead && thead.querySelectorAll("tr").length > 1;
		const hasCombinedCells = table.querySelector(
			"th[colspan], th[rowspan], td[colspan], td[rowspan]"
		);

		if (hasMultipleHeaders || hasCombinedCells) {
			const tableIndex = protectedTables.length;

			// Store the pristine HTML structure in the reference array for post-processing hydration
			protectedTables.push(table.outerHTML);

			// Inject a highly specific alphanumeric placeholder to bypass Defuddle table parsing
			const placeholder = doc.createElement("p");
			placeholder.textContent = `PROTECTEDTABLEPLACEHOLDER${tableIndex}`;
			table.replaceWith(placeholder);
		}
	});

	return doc.body.innerHTML;
}

function canHighlightElement(element: Element): boolean {
	// List of elements that cannot be nested inside mark tags
	const unsupportedElements = [
		"img",
		"video",
		"audio",
		"iframe",
		"canvas",
		"svg",
		"math",
		"table"
	];

	// Evaluate if the element contains any unsupported elements
	const hasUnsupportedElements = unsupportedElements.some(
		(tag) => element.getElementsByTagName(tag).length > 0
	);

	// Evaluate if the element itself is an unsupported type
	const isUnsupportedType = unsupportedElements.includes(
		element.tagName.toLowerCase()
	);

	return !hasUnsupportedElements && !isUnsupportedType;
}

function stripHtml(html: string): string {
	const parser = new DOMParser();
	const doc = parser.parseFromString(html, "text/html");
	return doc.body.textContent || "";
}

interface ContentResponse {
	author: string;
	content: string;
	description: string;
	domain: string;
	extractedContent: ExtractedContent;
	favicon: string;
	fullHtml: string;
	highlights: AnyHighlightData[];
	image: string;
	language: string;
	metaTags: {
		name?: string | null;
		property?: string | null;
		content: string | null;
	}[];
	parseTime: number;
	published: string;
	schemaOrgData: any;
	selectedHtml: string;
	site: string;
	title: string;
	wordCount: number;
}

async function sendExtractRequest(tabId: number): Promise<ContentResponse> {
	const response = (await browser.runtime.sendMessage({
		action: "sendMessageToTab",
		message: {
			action: "getPageContent"
		},
		tabId: tabId
	})) as ContentResponse & {
		success?: boolean;
		error?: string;
	};

	// Evaluate explicit errors returned from the background script
	if (
		response &&
		"success" in response &&
		!response.success &&
		response.error
	) {
		throw new Error(response.error);
	}

	if (response && response.content) {
		// Enforce strong typing on the returned highlight objects
		if (response.highlights && Array.isArray(response.highlights)) {
			response.highlights = response.highlights.map(
				(highlight: string | AnyHighlightData) => {
					if (typeof highlight === "string") {
						return {
							content: `<div>` + highlight + `</div>`,
							endOffset: highlight.length,
							id: Date.now().toString(),
							startOffset: 0,
							type: "text",
							xpath: ""
						};
					}
					return highlight as AnyHighlightData;
				}
			);
		} else {
			response.highlights = [];
		}
		return response;
	}
	throw new Error("No content received from page");
}

export async function extractPageContent(
	tabId: number
): Promise<ContentResponse | null> {
	try {
		return await sendExtractRequest(tabId);
	} catch (firstError) {
		// First attempt failed. This commonly happens on Safari after an
		// extension update when the old content script context is invalidated.
		// Retry execution.
		console.log(
			"[Obsidian Clipper] First extraction attempt failed, retrying...",
			firstError
		);
		try {
			return await sendExtractRequest(tabId);
		} catch (retryError) {
			console.error(
				"[Obsidian Clipper] Extraction failed after retry:",
				retryError
			);
			throw new Error(
				"Web Clipper was not able to start. Please try reloading the page."
			);
		}
	}
}

export async function initializePageContent(
	content: string,
	selectedHtml: string,
	extractedContent: ExtractedContent,
	currentUrl: string,
	schemaOrgData: any,
	fullHtml: string,
	highlights: AnyHighlightData[],
	title: string,
	author: string,
	description: string,
	favicon: string,
	image: string,
	published: string,
	site: string,
	wordCount: number,
	language: string,
	metaTags: {
		name?: string | null;
		property?: string | null;
		content: string | null;
	}[]
) {
	try {
		// Remove text fragments from URL
		currentUrl = currentUrl.replace(/#:~:text=[^&]+(&|$)/, "");

		// Step 1: Execute baseline figure and UI cleanup
		content = cleanAndFormatFigures(content);
		if (selectedHtml) {
			selectedHtml = cleanAndFormatFigures(selectedHtml);
		}

		// Step 2: Inject highlights into the live DOM before hiding tables
		if (
			generalSettings.highlighterEnabled &&
			generalSettings.highlightBehavior !== "no-highlights" &&
			highlights &&
			highlights.length > 0
		) {
			content = processHighlights(content, highlights);
			if (selectedHtml) {
				selectedHtml = processHighlights(selectedHtml, highlights);
			}
		}

		// Step 3: Extract complex tables into memory and inject placeholders
		const protectedTables: string[] = [];
		content = protectComplexTables(content, protectedTables);

		let selectedMarkdown = "";
		if (selectedHtml) {
			selectedHtml = protectComplexTables(selectedHtml, protectedTables);
			content = selectedHtml;
			selectedMarkdown = createMarkdownContent(selectedHtml, currentUrl);
		}

		// Step 4: Execute standard markdown conversion via Defuddle
		let markdownBody = createMarkdownContent(content, currentUrl);

		// Step 5: Hydrate the markdown with the pristine HTML tables
		protectedTables.forEach((tableHtml, index) => {
			const placeholder = `PROTECTEDTABLEPLACEHOLDER${index}`;
			const restoreHtml = `\n\n${tableHtml}\n\n`;

			// Replace globally in both standard and selected markdown outputs
			markdownBody = markdownBody.replace(
				new RegExp(placeholder, "g"),
				restoreHtml
			);
			if (selectedMarkdown) {
				selectedMarkdown = selectedMarkdown.replace(
					new RegExp(placeholder, "g"),
					restoreHtml
				);
			}
		});

		// Prepare highlight data for Obsidian properties
		const highlightsData = highlights.map((highlight) => {
			const highlightData: {
				text: string;
				timestamp: string;
				notes?: string[];
			} = {
				text: createMarkdownContent(highlight.content, currentUrl),
				timestamp: dayjs(parseInt(highlight.id)).toISOString()
			};
			if (highlight.notes && highlight.notes.length > 0) {
				highlightData.notes = highlight.notes;
			}
			return highlightData;
		});

		const noteName = sanitizeFileName(title);

		// Step 6: Build the final variables payload
		const currentVariables = buildVariables({
			author,
			content: markdownBody,
			contentHtml: content,
			description,
			extractedContent,
			favicon,
			fullHtml,
			highlights: highlights.length > 0 ? JSON.stringify(highlightsData) : "",
			image,
			language,
			metaTags,
			published,
			schemaOrgData,
			selection: selectedMarkdown,
			selectionHtml: selectedHtml,
			site,
			title,
			url: currentUrl,
			wordCount
		});

		debugLog("Variables", "Available variables:", currentVariables);

		return {
			currentVariables,
			noteName
		};
	} catch (error: unknown) {
		console.error("Error in initializePageContent:", error);
		if (error instanceof Error) {
			throw new Error(`Unable to initialize page content: ${error.message}`);
		} else {
			throw new Error("Unable to initialize page content: Unknown error");
		}
	}
}

function processHighlights(
	content: string,
	highlights: AnyHighlightData[]
): string {
	if (!generalSettings.highlighterEnabled || !highlights?.length) {
		return content;
	}

	if (generalSettings.highlightBehavior === "no-highlights") {
		return content;
	}

	if (generalSettings.highlightBehavior === "replace-content") {
		return highlights.map((highlight) => highlight.content).join("");
	}

	if (generalSettings.highlightBehavior === "highlight-inline") {
		debugLog("Highlights", "Using content length:", content.length);

		const parser = new DOMParser();
		const doc = parser.parseFromString(content, "text/html");
		const tempDiv = doc.body;

		const textHighlights = filterAndSortHighlights(highlights);
		debugLog("Highlights", "Processing highlights:", textHighlights.length);

		for (const highlight of textHighlights) {
			processHighlight(highlight, tempDiv as HTMLDivElement);
		}

		// Serialize the mutated DOM back to an HTML string
		const serializer = new XMLSerializer();
		let result = "";
		Array.from(tempDiv.childNodes).forEach((node) => {
			if (node.nodeType === Node.ELEMENT_NODE) {
				result += serializer.serializeToString(node);
			} else if (node.nodeType === Node.TEXT_NODE) {
				result += node.textContent;
			}
		});

		return result;
	}

	return content;
}

function filterAndSortHighlights(
	highlights: AnyHighlightData[]
): (TextHighlightData | ElementHighlightData)[] {
	return highlights
		.filter((h): h is TextHighlightData | ElementHighlightData => {
			if (h.type === "text") {
				return !!(h.xpath?.trim() || h.content?.trim());
			}
			if (h.type === "element" && h.xpath?.trim()) {
				const element = getElementByXPath(h.xpath);
				return element ? canHighlightElement(element) : false;
			}
			return false;
		})
		.sort((a, b) => {
			if (a.xpath && b.xpath) {
				const elementA = getElementByXPath(a.xpath);
				const elementB = getElementByXPath(b.xpath);
				if (elementA === elementB && a.type === "text" && b.type === "text") {
					return b.startOffset - a.startOffset;
				}
			}
			return 0;
		});
}

function processHighlight(
	highlight: TextHighlightData | ElementHighlightData,
	tempDiv: HTMLDivElement
) {
	try {
		if (highlight.xpath) {
			processXPathHighlight(highlight, tempDiv);
		} else {
			processContentBasedHighlight(highlight, tempDiv);
		}
	} catch (error) {
		debugLog("Highlights", "Error processing highlight:", error);
	}
}

function processXPathHighlight(
	highlight: TextHighlightData | ElementHighlightData,
	tempDiv: HTMLDivElement
) {
	const element = document.evaluate(
		highlight.xpath,
		tempDiv,
		null,
		XPathResult.FIRST_ORDERED_NODE_TYPE,
		null
	).singleNodeValue as Element;

	if (!element) {
		debugLog(
			"Highlights",
			"Could not find element for xpath:",
			highlight.xpath
		);
		return;
	}

	if (highlight.type === "element") {
		wrapElementWithMark(element);
	} else {
		wrapTextWithMark(element, highlight as TextHighlightData);
	}
}

function processContentBasedHighlight(
	highlight: TextHighlightData | ElementHighlightData,
	tempDiv: HTMLDivElement
) {
	const parser = new DOMParser();
	const doc = parser.parseFromString(highlight.content, "text/html");
	const contentDiv = doc.body;

	const serializer = new XMLSerializer();
	let innerContent = "";

	if (
		contentDiv.children.length === 1 &&
		contentDiv.firstElementChild?.tagName === "DIV"
	) {
		Array.from(contentDiv.firstElementChild.childNodes).forEach((node) => {
			if (node.nodeType === Node.ELEMENT_NODE) {
				innerContent += serializer.serializeToString(node);
			} else if (node.nodeType === Node.TEXT_NODE) {
				innerContent += node.textContent;
			}
		});
	} else {
		Array.from(contentDiv.childNodes).forEach((node) => {
			if (node.nodeType === Node.ELEMENT_NODE) {
				innerContent += serializer.serializeToString(node);
			} else if (node.nodeType === Node.TEXT_NODE) {
				innerContent += node.textContent;
			}
		});
	}

	const paragraphs = Array.from(contentDiv.querySelectorAll("p"));
	if (paragraphs.length) {
		processContentParagraphs(paragraphs, tempDiv);
	} else {
		processInlineContent(innerContent, tempDiv);
	}
}

function processContentParagraphs(
	sourceParagraphs: Element[],
	tempDiv: HTMLDivElement
) {
	sourceParagraphs.forEach((sourceParagraph) => {
		const sourceText = stripHtml(sourceParagraph.outerHTML).trim();
		debugLog("Highlights", "Looking for paragraph:", sourceText);

		const paragraphs = Array.from(tempDiv.querySelectorAll("p"));
		for (const targetParagraph of paragraphs) {
			const targetText = stripHtml(targetParagraph.outerHTML).trim();

			if (targetText === sourceText) {
				debugLog(
					"Highlights",
					"Found matching paragraph:",
					targetParagraph.outerHTML
				);
				wrapElementWithMark(targetParagraph);
				break;
			}
		}
	});
}

function processInlineContent(content: string, tempDiv: HTMLDivElement) {
	const searchText = stripHtml(content).trim();
	debugLog("Highlights", "Searching for text:", searchText);

	const walker = document.createTreeWalker(tempDiv, NodeFilter.SHOW_TEXT);

	let node;
	while ((node = walker.nextNode() as Text)) {
		const nodeText = node.textContent || "";
		const index = nodeText.indexOf(searchText);

		if (index !== -1) {
			debugLog("Highlights", "Found matching text in node:", {
				index: index,
				text: nodeText
			});

			const range = document.createRange();
			range.setStart(node, index);
			range.setEnd(node, index + searchText.length);

			const mark = document.createElement("mark");
			range.surroundContents(mark);
			debugLog("Highlights", "Created mark element:", mark.outerHTML);
			break;
		}
	}
}
