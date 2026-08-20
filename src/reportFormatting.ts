export interface ReportSource {
  title: string;
  url: string;
}

const REFERENCE_HEADING = "## 五、参考网站 (References)";
const NO_REFERENCE_TEXT = "未检索到可公开核验的网站来源。";

const cleanMarkdownLabel = (line: string, url: string): string => line
  .replace(/^\s*(?:#{1,6}\s*)?[-*+]?\s*/, "")
  .replace(/\*\*/g, "")
  .replace(/^(?:五[、.．]\s*)?(?:参考网站|参考来源|网络参考|References?)\s*[:：]?\s*/i, "")
  .replace(url, "")
  .replace(/^\[|\]\s*\(?$/g, "")
  .replace(/^[：:：\-–—\s]+|[：:：\-–—\s]+$/g, "")
  .trim();

const extractReferences = (line: string): ReportSource[] => {
  const references: ReportSource[] = [];
  const markdownLinkPattern = /\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/gi;
  let markdownMatch: RegExpExecArray | null;
  const markdownUrls = new Set<string>();
  while ((markdownMatch = markdownLinkPattern.exec(line)) !== null) {
    const url = markdownMatch[2].replace(/[.,，。;；:：]+$/g, "");
    markdownUrls.add(url);
    references.push({ title: markdownMatch[1].trim() || url, url });
  }

  const plainUrlPattern = /https?:\/\/[^\s<>{}\[\]]+/gi;
  let urlMatch: RegExpExecArray | null;
  while ((urlMatch = plainUrlPattern.exec(line)) !== null) {
    const url = urlMatch[0].replace(/[)\].,，。;；:：]+$/g, "");
    if (markdownUrls.has(url)) continue;
    references.push({ title: cleanMarkdownLabel(line, url) || url, url });
  }
  return references;
};

const normalizedMarkerText = (line: string): string => line
  .replace(/^\s*(?:#{1,6}\s*)?[-*+]?\s*/, "")
  .replace(/\*\*/g, "")
  .trim();

const isReferenceStart = (line: string): boolean =>
  /^(?:五[、.．]\s*)?(?:参考网站|参考来源|网络参考|References?)\s*[:：]?/i.test(normalizedMarkerText(line));

const isOrderInfo = (line: string): boolean =>
  /^(?:订单收货信息|订单信息)\s*[:：]?/i.test(normalizedMarkerText(line));

const isSectionBoundary = (line: string): boolean => {
  const trimmed = line.trim();
  return /^#{1,6}\s+/.test(trimmed) || /^(?:一|二|三|四|五|六|七|八|九|十)[、.．]\s*/.test(normalizedMarkerText(line));
};

const addUniqueReference = (target: ReportSource[], seenUrls: Set<string>, source: ReportSource) => {
  const url = String(source.url || "").trim().replace(/[)\].,，。;；:：]+$/g, "");
  if (!url || seenUrls.has(url)) return;
  seenUrls.add(url);
  target.push({ title: String(source.title || "").trim() || url, url });
};

export function normalizeReportReferenceSection(report: string, fallbackSources: ReportSource[] = []): string {
  const lines = String(report || "").replace(/\r\n?/g, "\n").split("\n");
  const bodyLines: string[] = [];
  const references: ReportSource[] = [];
  const seenUrls = new Set<string>();

  for (let index = 0; index < lines.length;) {
    if (!isReferenceStart(lines[index])) {
      bodyLines.push(lines[index]);
      index++;
      continue;
    }

    extractReferences(lines[index]).forEach(source => addUniqueReference(references, seenUrls, source));
    index++;

    while (index < lines.length) {
      const currentLine = lines[index];
      if (isOrderInfo(currentLine) || (isSectionBoundary(currentLine) && !isReferenceStart(currentLine))) {
        break;
      }
      extractReferences(currentLine).forEach(source => addUniqueReference(references, seenUrls, source));
      index++;
    }

    while (bodyLines.length > 0 && bodyLines[bodyLines.length - 1].trim() === "") bodyLines.pop();
  }

  if (references.length === 0) {
    fallbackSources.forEach(source => addUniqueReference(references, seenUrls, source));
  }

  const body = bodyLines.join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
  const referenceLines = references.length > 0
    ? references.map(source => `- ${source.title}：${source.url}`)
    : [NO_REFERENCE_TEXT];

  return `${body}\n\n${REFERENCE_HEADING}\n${referenceLines.join("\n")}`.trim();
}
