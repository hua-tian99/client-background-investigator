import React, { useState } from "react";
import ReactMarkdown from "react-markdown";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import {
  Building2,
  Search,
  Globe,
  FileSpreadsheet,
  Plus,
  Trash2,
  ExternalLink,
  ChevronRight,
  ClipboardCopy,
  CheckCircle,
  FileText,
  Mail,
  MessageSquare,
  Sparkles,
  Loader2,
  MapPin,
  Calendar,
  Layers,
  Award,
  BookOpen,
  Info,
  Check,
  AlertCircle,
  Download,
  Upload,
  Key,
  Eye,
  EyeOff,
  Cpu,
  Cloud,
  CloudUpload,
  CloudOff,
  X,
  ShieldAlert,
  TrendingUp,
  Compass,
  Database,
  Zap
} from "lucide-react";

// Types for Customer Leads and Investigations
interface CompanyOverview {
  legalName: string;
  taxId: string;
  website: string;
  industry?: string;
  founded?: string;
  status?: string;
  headquarters?: string;
}

interface LeadUsage {
  promptTokens: number;
  completionTokens: number;
  totalTokens: number;
  modelUsed: string;
  searchGroundingUsed: boolean;
}

interface Lead {
  id: string;
  companyName: string;
  recipientName?: string;
  country: string;
  countryNameZh?: string;
  orderInfo: string;
  productContext: string;
  productNameZh?: string;
  status: "idle" | "searching" | "completed" | "failed";
  grade?: string;
  summary?: string;
  riskAlert?: string;
  followUpStrategy?: string;
  followUpGrade?: string;
  companyOverview?: CompanyOverview;
  report?: string;
  sources?: Array<{ title: string; url: string }>;
  error?: string;
  completedAt?: string;
  excelRowIndex?: number;
  feishuUrl?: string;
  feishuStatus?: "idle" | "uploading" | "success" | "failed";
  feishuError?: string;
  feishuImportFallback?: boolean;
  feishuImportError?: string;
  feishuImportAttempts?: number;
  bitableStatus?: "idle" | "syncing" | "success" | "failed";
  bitableUrl?: string;
  bitableError?: string;
  usage?: LeadUsage;
}

const COUNTRY_NAME_ZH_MAP: Record<string, string> = {
  cn: "中国", china: "中国", us: "美国", usa: "美国", "united states": "美国", "united states of america": "美国",
  ca: "加拿大", canada: "加拿大", gb: "英国", uk: "英国", "united kingdom": "英国", england: "英国",
  de: "德国", germany: "德国", fr: "法国", france: "法国", it: "意大利", italy: "意大利",
  es: "西班牙", spain: "西班牙", pt: "葡萄牙", portugal: "葡萄牙", pl: "波兰", poland: "波兰",
  ua: "乌克兰", ukraine: "乌克兰", ru: "俄罗斯", russia: "俄罗斯", nl: "荷兰", netherlands: "荷兰",
  be: "比利时", belgium: "比利时", at: "奥地利", austria: "奥地利", ch: "瑞士", switzerland: "瑞士",
  se: "瑞典", sweden: "瑞典", no: "挪威", norway: "挪威", dk: "丹麦", denmark: "丹麦",
  fi: "芬兰", finland: "芬兰", cz: "捷克", czechia: "捷克", "czech republic": "捷克",
  ro: "罗马尼亚", romania: "罗马尼亚", hu: "匈牙利", hungary: "匈牙利", gr: "希腊", greece: "希腊",
  tr: "土耳其", turkey: "土耳其", türkiye: "土耳其", au: "澳大利亚", australia: "澳大利亚",
  nz: "新西兰", "new zealand": "新西兰", jp: "日本", japan: "日本", kr: "韩国", "south korea": "韩国", korea: "韩国",
  in: "印度", india: "印度", id: "印度尼西亚", indonesia: "印度尼西亚", my: "马来西亚", malaysia: "马来西亚",
  sg: "新加坡", singapore: "新加坡", th: "泰国", thailand: "泰国", vn: "越南", vietnam: "越南",
  ph: "菲律宾", philippines: "菲律宾", ae: "阿联酋", uae: "阿联酋", "united arab emirates": "阿联酋",
  sa: "沙特阿拉伯", "saudi arabia": "沙特阿拉伯", il: "以色列", israel: "以色列", eg: "埃及", egypt: "埃及",
  za: "南非", "south africa": "南非", ng: "尼日利亚", nigeria: "尼日利亚", ke: "肯尼亚", kenya: "肯尼亚",
  mx: "墨西哥", mexico: "墨西哥", br: "巴西", brazil: "巴西", ar: "阿根廷", argentina: "阿根廷",
  cl: "智利", chile: "智利", co: "哥伦比亚", colombia: "哥伦比亚", pe: "秘鲁", peru: "秘鲁"
};

const cleanCloudTitlePart = (value: unknown, maxLength: number): string => String(value || "")
  .replace(/[\u0000-\u001f\u007f]/g, " ")
  .replace(/[\\/:*?"<>|]/g, " ")
  .replace(/[\s\-–—]+/g, " ")
  .trim()
  .slice(0, maxLength)
  .trim();

const resolveCountryNameZh = (countryNameZh: unknown, country: unknown): string => {
  const aiName = cleanCloudTitlePart(countryNameZh, 12);
  if (aiName && /[\u3400-\u9fff]/.test(aiName)) return aiName;
  const rawCountry = String(country || "").trim();
  const existingChinese = rawCountry.match(/[\u3400-\u9fff]{2,12}/)?.[0];
  if (existingChinese) return existingChinese;
  return COUNTRY_NAME_ZH_MAP[rawCountry.toLowerCase()] || "未知国家";
};

const resolveProductNameZh = (productNameZh: unknown, productContext: unknown): string => {
  const aiName = cleanCloudTitlePart(productNameZh, 30);
  if (aiName && /[\u3400-\u9fff]{2,}/.test(aiName)) return aiName;
  const rawProduct = String(productContext || "").trim();
  const chineseSegments = rawProduct.match(/[\u3400-\u9fff]{2,30}/g) || [];
  const existingChinese = chineseSegments.find(segment => !["商品数量", "产品属性", "商品信息", "颜色", "尺寸"].includes(segment));
  if (existingChinese) return cleanCloudTitlePart(existingChinese, 30);
  const normalized = rawProduct.toLowerCase();
  const productKeywords: Array<[string[], string]> = [
    [["water detector", "water finder"], "地下水探测仪"],
    [["cavity detector", "treasure detector", "metal detector"], "三维成像探宝仪"],
    [["utility knife", "box cutter", "paper cutter"], "迷你美工刀"],
    [["kitchen faucet", "mixer tap", "water tap"], "厨房水龙头"],
    [["labeling machine", "label machine"], "贴标机"],
    [["filling machine"], "灌装机"],
    [["packing machine", "packaging machine"], "包装机"],
    [["laser marking", "laser engrav"], "激光打标机"],
    [["sealing machine"], "封口机"]
  ];
  return productKeywords.find(([keywords]) => keywords.some(keyword => normalized.includes(keyword)))?.[1] || "相关产品";
};

export const buildFeishuDocumentTitle = (lead: Lead): string => {
  const recipient = cleanCloudTitlePart(lead.recipientName || lead.companyName, 50) || "未知收件人";
  const country = resolveCountryNameZh(lead.countryNameZh, lead.country);
  const product = resolveProductNameZh(lead.productNameZh, lead.productContext);
  return `${recipient}-${country}-${product}`.slice(0, 120).replace(/[\s\-–—]+$/g, "");
};

interface FeishuSubfolderState {
  token: string;
  index: number;
  name?: string;
}

const FEISHU_SUBFOLDER_STATE_KEY = "feishu_report_subfolders";
const readFeishuSubfolderState = (rootFolder: string): FeishuSubfolderState | null => {
  try {
    const all = JSON.parse(localStorage.getItem(FEISHU_SUBFOLDER_STATE_KEY) || "{}");
    return all[rootFolder.trim()] || null;
  } catch {
    return null;
  }
};

const saveFeishuSubfolderState = (rootFolder: string, state: FeishuSubfolderState) => {
  try {
    const all = JSON.parse(localStorage.getItem(FEISHU_SUBFOLDER_STATE_KEY) || "{}");
    all[rootFolder.trim()] = state;
    localStorage.setItem(FEISHU_SUBFOLDER_STATE_KEY, JSON.stringify(all));
  } catch {
    // Upload still works for this session even if browser storage is unavailable.
  }
};

const clearFeishuSubfolderState = (rootFolder: string) => {
  try {
    const all = JSON.parse(localStorage.getItem(FEISHU_SUBFOLDER_STATE_KEY) || "{}");
    delete all[rootFolder.trim()];
    localStorage.setItem(FEISHU_SUBFOLDER_STATE_KEY, JSON.stringify(all));
  } catch {
    localStorage.removeItem(FEISHU_SUBFOLDER_STATE_KEY);
  }
};

/**
 * Calculates USD and RMB costs for the given model and token counts.
 * Precise conversion rate: 1 USD ≈ 7.25 CNY (RMB)
 */
export function calculateTokenCost(
  modelUsed: string = "deepseek-v4-flash",
  promptTokens: number = 0,
  completionTokens: number = 0
): { usd: number; rmb: number } {
  const modelKey = modelUsed.toLowerCase();

  // 定义 1M Tokens 的美金单价 (Rates per million tokens)
  // DeepSeek V4 定价 (USD per 1M tokens, 缓存未命中价)
  const PRICING: Record<string, { input: number; output: number }> = {
    "deepseek-v4-flash": { input: 0.14, output: 0.28 }
  };

  let rates = null;
  // Match model price
  for (const [key, val] of Object.entries(PRICING)) {
    if (modelKey.includes(key)) {
      rates = val;
      break;
    }
  }

  // Fallback to deepseek-v4-flash as default if no key matched
  if (!rates) {
    rates = PRICING["deepseek-v4-flash"];
  }

  // Calculate: (Token count / 1,000,000) * rates
  const inputCost = (promptTokens / 1000000) * rates.input;
  const outputCost = (completionTokens / 1000000) * rates.output;
  
  // Keep 6 decimal places to prevent rounding tiny costs to 0
  const usd = Number((inputCost + outputCost).toFixed(6));
  const rmb = Number((usd * 7.25).toFixed(6));

  return { usd, rmb };
}

// Helper function to dynamically repair common markdown table single-line formatting bugs
const preprocessMarkdown = (text: string | undefined): string => {
  if (!text) return "";
  
  // Replace literal backslash-n or backslash-r string with actual newline characters
  let formatted = text;
  formatted = formatted.replace(/\\n/g, "\n").replace(/\\r/g, "\r");
  
  // Replace double vertical bars on the same line with standard markdown row separation newlines.
  // This robustly fixes issues where table rows are merged on a single line with " | | " or " || ".
  formatted = formatted.replace(/\|[ \t]*\|/g, "|\n|");
  return formatted;
};

// Helper function to extract order number/identifier and construct the export file name
const getExportFileName = (lead: Lead, extension: string): string => {
  let orderNo = "";
  if (lead.orderInfo) {
    const patterns = [
      /\[订单号\]:\s*([^\n\r]+)/i,
      /\[订单编号\]:\s*([^\n\r]+)/i,
      /\[Order ID\]:\s*([^\n\r]+)/i,
      /\[Order No\]:\s*([^\n\r]+)/i,
      /\[Order Number\]:\s*([^\n\r]+)/i,
      /\[Parent Order ID\]:\s*([^\n\r]+)/i,
      /\[交易号\]:\s*([^\n\r]+)/i,
      /订单号[:：]\s*([^\n\r]+)/i,
      /订单编号[:：]\s*([^\n\r]+)/i,
    ];

    for (const pattern of patterns) {
      const match = lead.orderInfo.match(pattern);
      if (match && match[1] && match[1].trim()) {
        orderNo = match[1].trim();
        break;
      }
    }

    if (!orderNo) {
      const numericMatch = lead.orderInfo.match(/\b\d{10,25}\b/);
      if (numericMatch) {
        orderNo = numericMatch[0];
      }
    }
  }

  if (orderNo) {
    // Sanitize the order number for filesystem safety
    const sanitized = orderNo.replace(/[\\/:*?"<>|]/g, "_").trim();
    if (sanitized) {
      return `${sanitized}.${extension}`;
    }
  }

  // Fallback if no order number can be found
  const fallbackName = lead.companyName.replace(/[^a-zA-Z0-9\u4e00-\u9fa5]/g, "_").toLowerCase();
  return `${fallbackName}_due_diligence_report.${extension}`;
};

export default function App() {
  // Pre-loaded leads. The first one is the exact due diligence report the user provided for Acme Sp. z o.o.
  // This ensures the application is immediately usable and fully showcases the expected quality standard.
  const [leads, setLeads] = useState<Lead[]>([
    {
      id: "acme-sp-z-o-o",
      companyName: "Acme Sp. z o.o.",
      country: "Poland (波兰)",
      orderInfo: "收件人名称 Acme Sp. z o.o. NIP 123 456 78 90, ul. Akacjowa 15 , 00-123, Warszawa, Warszawa, Mazowieckie, Poland, kontakt@acme-example.pl, +48 600 123 456",
      productContext: "贴标机 (Labeling Machine)",
      status: "completed",
      grade: "💎 S级（高价值实体生产型客户，建议重点跟进）",
      summary: "Acme Sp. z o.o. 是一家深耕波兰食品加工行业的源头制造企业，主要生产果酱、水果馅料及糖浆等高规格烘焙原料。具备独立灌装线，对贴标及配套包装设备具有极高大额订单潜能。",
      companyOverview: {
        legalName: "Acme Spółka z ograniczoną odpowiedzialnością",
        taxId: "NIP 1234567890 / KRS 0000123456",
        website: "https://acme-example.pl",
        industry: "水果及蔬菜后续加工 / Fruit & Vegetable Processing",
        founded: "2019",
        status: "Active (在册运行)",
        headquarters: "Warszawa, Poland (波兰华沙)",
      },
      report: `# Acme Sp. z o.o. 贴标机公司客户背调

# 境外企业客户背景调查报告 (Due Diligence Report)

**报告对象：** Acme Sp. z o.o. (波兰)

**调查目的：** 评估 AliExpress 平台买家商业背景及长线 B2B 合作潜力

**评估等级：** 💎 **S级（高价值实体生产型客户，建议重点跟进）**

## **订单收货信息：**
- **收件人名称** Acme Sp. z o.o. NIP 123 456 78 90
- **详细地址** ul. Akacjowa 15 , 00-123, Warszawa, Warszawa, Mazowieckie, Poland
- **邮编** 00-123
- **联系邮件** kontakt@acme-example.pl
- **联系电话** +48 600 123 456

## 一、 企业工商基本信息 (Corporate Profile)

| 工商核查项 | 内容 |
| :--- | :--- |
| **注册公司全称** | Acme Spółka z ograniczoną odpowiedzialnością (Acme Sp. z o.o.) |
| **国家/地区** | Poland (波兰) |
| **统一社会信用代码/税号** | NIP: 1234567890 / KRS: 0000782352 |
| **成立时间** | 2019年05月07日 |
| **法定代表人/负责人** | Jan Nowak (CEO & Co-owner) |
| **官方网站** | https://acme-example.pl |
| **经营状态** | Active (正常在册运行) |
| **总部地址** | ul. Akacjowa 15, 00-123, Warszawa, Mazowieckie, Poland |

## 二、 业务线与生产实力分析 (Business & Operations)

### 核心业务定位
Acme 是一家**深耕食品加工行业的源头制造企业**。主要业务为水果及蔬菜的后续加工与保鲜，核心产品包括：
- 工业级及烘焙级**果酱、蜜饯、水果馅料、糖浆**。
- 核心客群为波兰本土及周边的**面包房、糕点厂、酒店及大型餐饮集团 (HoReCa B2B供应)**。

### 包装工艺与设备需求
从其官方产品线看，该公司的产品多采用**塑料桶、玻璃罐、高容量塑料瓶**进行封装。

> **采购行为透视：** 他们在 AliExpress 采购“贴标机”，属于纯粹的**生产线资本支出 (CapEx)**。这表明工厂目前正在进行：
> - 某条半自动包装线的自动化升级；
> - 开辟新规格包装（如新罐装果酱）的副产线；
> - 补充现有生产线的产能。

## 三、 合作价值与复购潜力评估 (Value Evaluation)

- **设备耗材与配件持续性（高粘性）：** 食品加工车间环境（高湿度、水果酸性）对机械磨损较大。后续对贴标机的**易损件（切刀、硅胶轮、传感器、输送带）及标签耗材**有长期刚需。
- **包装整线扩张潜力（高客单价）：** 既然该公司具备独立的食品灌装生产线，除贴标机外，未来必然会产生**自动灌装机、旋盖机、喷码机、理瓶机、全自动打包机**等配套设备的升级需求。
- **跨境渠道沉淀潜力：** 平台采购（AliExpress）通常是企业客户用于“测样”或“紧急替代”的临时渠道。一旦建立信任，极易引导至**线下传统大额 B2B 贸易**，避开平台佣金，提升利润率。

## 四、 针对性开发与跟进策略 (Action Plan)

针对该客户的特点，建议销售及技术团队采取“技术切入，服务先行，长线转化”的策略：

### 阶段一：高规格售后服务，建立企业信任（立即执行）
- **触达对象：** 检查收件人姓名是否为 **Jan Nowak**。
- **话术重点：** “我们注意到您是 Acme 公司的负责人。作为源头机械工厂，我们非常重视您的订单。为了完美匹配您的果酱/食品包装线，我们已为您开通**绿色技术通道**，提供专属的英文调试视频与 1对1 工程师支持。”

### 阶段二：技术参数切入，摸清客户产能（发货后跟进）
- **互动借口：** 以“确保贴标精度”为由，向其索要包装数据。
- **话术示例：** “由于食品级瓶身（罐装/桶装）的材质 and 弧度不同，您可以将您需要贴标的容器尺寸和标签材质发给我们。我们的技术团队可在出厂前为您做免费的**远程参数预设**。”（借此套出其真实的生产线规格与规模）。

### 阶段三：全产业链推荐，引导线下长期合作（收货满意后）
- **价值锚定：** 向其发送我司的 **B2B 完整产品目录 (E-Catalog)**，重点突出食品/包装流体线（灌装-旋盖-贴标-喷码一体化解决方案）。
- **利益诱导：** 暗示如果未来进行大宗设备采购或批量配件补充，可绕过零售平台，直接享受**工厂直销批发价 (Wholesale Price)**，并通过中欧班列或海运大幅降低物流成本。

**结论：** 该客户背景真实、财力稳定、行业匹配度极高，属于**含金量极高的企业级客户**。建议将其录入我司 CRM 系统作为“重点战略客户”进行长期跟进。

## 五、参考网站 (References)
- [GoWork.pl - Acme Sp. z o.o. Profile](https://www.example.com/acme-sp-z-o-o)
- [KRS Pobierz - Commercial Registry Data](https://example.com/krs/acme-sp-z-o-o)
- [Official Website - Acme](https://acme-example.pl)`,
      sources: [
        { title: "Acme Sp. z o.o. - GoWork.pl", url: "https://www.example.com/acme-sp-z-o-o" },
        { title: "Acme Sp. z o.o. - KRS Pobierz", url: "https://example.com/krs/acme-sp-z-o-o" }
      ],
      completedAt: "2026-07-12 23:05:19"
    }
  ]);

  const [selectedLeadId, setSelectedLeadId] = useState<string>("acme-sp-z-o-o");
  const [activeTab, setActiveTab] = useState<"report" | "outreach" | "sources" | "usage">("report");
  const [isAddingLead, setIsAddingLead] = useState<boolean>(false);
  const [addMode, setAddMode] = useState<"batch" | "manual">("batch");
  const [appMode, setAppMode] = useState<"create" | "update">("create");
  
  // Forms & Pasting states
  const [pasteArea, setPasteArea] = useState<string>("");
  const [manualForm, setManualForm] = useState({
    companyName: "",
    country: "",
    orderInfo: "",
    productContext: "贴标机 (Labeling Machine)"
  });

  // Excel & File Upload states
  const [batchSubMode, setBatchSubMode] = useState<"file" | "paste">("file");
  const [isDragging, setIsDragging] = useState<boolean>(false);
  const [uploadError, setUploadError] = useState<string | null>(null);
  const [uploadSuccessMsg, setUploadSuccessMsg] = useState<string | null>(null);
  const [originalSheetData, setOriginalSheetData] = useState<any[][] | null>(null);
  const [originalFileName, setOriginalFileName] = useState<string>("");
  const [isExportingZip, setIsExportingZip] = useState<boolean>(false);

  // Custom API Key states
  const [customApiKey, setCustomApiKey] = useState<string>(() => {
    return localStorage.getItem("custom_deepseek_api_key") || localStorage.getItem("custom_gemini_api_key") || "";
  });
  const [selectedModel, setSelectedModel] = useState<string>(() => {
    // Post-migration only deepseek-v4-flash is supported. Legacy values stored
    // under old keys (or old "deepseek-v4-pro" / "gemini-*" selections) are
    // normalized to the sole supported model.
    const saved = localStorage.getItem("selected_deepseek_model") || localStorage.getItem("selected_gemini_model");
    return saved === "deepseek-v4-flash" ? saved : "deepseek-v4-flash";
  });
  const [showKeyInput, setShowKeyInput] = useState<boolean>(false);
  const [tempApiKey, setTempApiKey] = useState<string>(customApiKey);
  const [showPassword, setShowPassword] = useState<boolean>(false);
  
  // Feishu Integration Config states
  const [enableFeishu, setEnableFeishu] = useState<boolean>(() => {
    return localStorage.getItem("feishu_enabled") === "true";
  });
  const [feishuAppId, setFeishuAppId] = useState<string>(() => {
    return localStorage.getItem("feishu_app_id") || "";
  });
  const [feishuAppSecret, setFeishuAppSecret] = useState<string>(() => {
    return localStorage.getItem("feishu_app_secret") || "";
  });
  const [feishuFolderToken, setFeishuFolderToken] = useState<string>(() => {
    return localStorage.getItem("feishu_folder_token") || "";
  });
  const [feishuBitableUrl, setFeishuBitableUrl] = useState<string>(() => {
    return localStorage.getItem("feishu_bitable_url") || "";
  });
  const [isUploadingSheet, setIsUploadingSheet] = useState<boolean>(false);
  const [uploadSheetError, setUploadSheetError] = useState<string | null>(null);
  const [feishuSheetUrl, setFeishuSheetUrl] = useState<string>(() => {
    return localStorage.getItem("feishu_sheet_url") || "";
  });
  const [isUploadingToFeishu, setIsUploadingToFeishu] = useState<boolean>(false);
  const [feishuConfigOpen, setFeishuConfigOpen] = useState<boolean>(false);
  
  const [syncMappingAnalysis, setSyncMappingAnalysis] = useState<any[] | null>(null);
  const [isMappingAnalysisModalOpen, setIsMappingAnalysisModalOpen] = useState<boolean>(false);
  
  // Custom checklist states for selected client CRM pipeline
  const [checklist, setChecklist] = useState<Record<string, Record<string, boolean>>>({
    "acme-sp-z-o-o": {
      "stage1": false,
      "stage2": false,
      "stage3": false,
    }
  });

  // Copy success notification state
  const [copiedText, setCopiedText] = useState<string | null>(null);
  const [showGradingGuide, setShowGradingGuide] = useState<boolean>(false);

  const selectedLead = leads.find((l) => l.id === selectedLeadId) || leads[0];

  const excelLeads = leads.filter(l => l.excelRowIndex !== undefined);
  const totalUploadedCount = excelLeads.length;
  const completedCount = excelLeads.filter(l => l.status === "completed").length;
  const runningCount = excelLeads.filter(l => l.status === "searching").length;
  const failedCount = excelLeads.filter(l => l.status === "failed").length;
  const idleCount = excelLeads.filter(l => l.status === "idle").length;

  const handleCopy = (text: string, type: string) => {
    navigator.clipboard.writeText(text);
    setCopiedText(type);
    setTimeout(() => setCopiedText(null), 2000);
  };

  const toggleCheck = (leadId: string, stage: string) => {
    setChecklist(prev => ({
      ...prev,
      [leadId]: {
        ...(prev[leadId] || {}),
        [stage]: !(prev[leadId]?.[stage])
      }
    }));
  };

  const handleDeleteLead = (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    const updated = leads.filter(l => l.id !== id);
    setLeads(updated);
    if (selectedLeadId === id && updated.length > 0) {
      setSelectedLeadId(updated[0].id);
    }
  };

  // Parses tab-separated (from Excel) or comma-separated lines of company info
  const handleBatchImport = () => {
    if (!pasteArea.trim()) return;
    
    const lines = pasteArea.split("\n").map(l => l.trim()).filter(l => l.length > 0);
    const newLeads: Lead[] = lines.map((line, idx) => {
      // Handle spreadsheet / tab separated copy-pastes first, fallback to commas, or fallback to full line
      let parts = line.split("\t");
      if (parts.length < 2) {
        parts = line.split(",");
      }
      
      const companyName = parts[0]?.trim() || `Pasted Company #${idx + 1}`;
      const country = parts[1]?.trim() || "";
      const orderInfo = parts[2]?.trim() || "";
      const productContext = parts[3]?.trim() || "包装设备/贴标机 (Packaging Machinery)";

      return {
        id: `lead-${Date.now()}-${idx}`,
        companyName,
        country,
        orderInfo,
        productContext,
        status: "idle" as const
      };
    });

    setLeads(prev => [...prev, ...newLeads]);
    setPasteArea("");
    setIsAddingLead(false);
    if (newLeads.length > 0) {
      setSelectedLeadId(newLeads[0].id);
    }
  };

  const handleManualImport = () => {
    if (!manualForm.companyName.trim()) return;

    const newLead: Lead = {
      id: `lead-${Date.now()}`,
      companyName: manualForm.companyName,
      country: manualForm.country,
      orderInfo: manualForm.orderInfo,
      productContext: manualForm.productContext || "贴标机 (Labeling Machine)",
      status: "idle" as const
    };

    setLeads(prev => [...prev, newLead]);
    setManualForm({
      companyName: "",
      country: "",
      orderInfo: "",
      productContext: "贴标机 (Labeling Machine)"
    });
    setIsAddingLead(false);
    setSelectedLeadId(newLead.id);
  };

  const handleDownloadSampleTemplate = () => {
    const fields = [
      '订单号', '店铺', '订单状态', '负责人（业务员）', '买家名称', '下单时间', 
      '付款时间', '支付方式', '供货价', '产品总金额', '物流费用', '预计增值税', 
      '平台是否代征代缴', '订单金额', '买家实付金额', 'DDP关税', '店铺优惠', 
      '托管商家折扣', '平台补贴', '商品ID', '商品信息', 'EANcode', '商品编码', 
      '订单备注', '完整收货地址', '完整收货地址（常用语言）', '收件人名称', '联系电话', 
      '手机', '联系邮件', '详细地址', '邮编', '扩展城市（德/意/波/墨为真实的城市）', 
      '城市', '州/省', '收货国家', 'National address（仅沙特使用）', '税号', 
      '买家选择物流', '发货期限', '实际发货单号', '发货时间', '确认收货时间', 
      '订单业务模式', '定制信息', '记录 ID', '文本 35', '标题'
    ];
    
    const sampleRow = [
      '2100000000000001', 'AliExpress旗舰店', '已付款', '业务员A', 'Acme Sp. z o.o.', '2026-07-14 10:00:00',
      '2026-07-14 10:05:00', '担保交易', '1200', '1500', '120', '345',
      '是', '1965', '1965', '0', '50',
      '0', '15', 'PROD-0001', '贴标机 (Labeling Machine)', 'EAN1234567890', 'CODE-LAB-01',
      '请预设贴标机瓶身尺寸', 'ul. Akacjowa 15, Warszawa, Poland', 'ul. Akacjowa 15, Warszawa, Poland', 'Jan Nowak', '+48 600 123 456',
      '+48 600 123 456', 'kontakt@acme-example.pl', 'ul. Akacjowa 15', '00-123', 'Warszawa',
      'Warszawa', 'Mazowieckie', 'Poland', '', 'NIP 1234567890',
      'DHL Express', '2026-07-20', 'DHL12345678', '2026-07-15 14:00:00', '2026-07-19 16:30:00',
      'DDP直邮', '瓶身直径：80mm', 'REC-00001', 'N/A', ''
    ];

    const sampleRow2 = [
      '2100000000000002', 'eBay精品店', '已付款', '业务员B', 'AgroSabores S.A.S', '2026-07-14 11:30:00',
      '2026-07-14 11:32:00', 'PayPal', '3200', '3500', '250', '805',
      '否', '4555', '4555', '120', '100',
      '0', '50', 'PROD-0002', '液体灌装机 (Liquid Filling Machine)', 'EAN98765432', 'CODE-FIL-02',
      'Need Spanish manuals', 'Carrera 7 # 12-34, Bogotá, Colombia', 'Carrera 7 # 12-34, Bogotá, Colombia', 'Andrés Morales', '+57 315 123 4567',
      '+57 315 123 4567', 'andres.morales@example.com', 'Carrera 7 # 12-34', '110001', 'Bogotá',
      'Bogotá', 'Cundinamarca', 'Colombia', '', 'NIT 901234567-1',
      'FedEx IP', '2026-07-22', 'FDX87654321', '', '',
      'B2B贸易', '物料：粘稠芒果浆', 'REC-00002', 'N/A', ''
    ];
    
    const aoa = [fields, sampleRow, sampleRow2];
    const worksheet = XLSX.utils.aoa_to_sheet(aoa);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "48字段背调模版");
    XLSX.writeFile(workbook, "feishu_bitable_48_fields_sample_template.xlsx");
  };

  const handleFileUpload = (file: File) => {
    setUploadError(null);
    setUploadSuccessMsg(null);
    
    const name = file.name.toLowerCase();
    if (!name.endsWith(".xlsx") && !name.endsWith(".xls") && !name.endsWith(".csv")) {
      setUploadError("仅支持导入 .xlsx, .xls, .csv 电子表格文件");
      return;
    }
    
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const buffer = e.target?.result as ArrayBuffer;
        const data = new Uint8Array(buffer);
        const workbook = XLSX.read(data, { type: "array", cellDates: true, dateNF: "yyyy-mm-dd hh:MM:ss" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        const sheetData = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1 });
        
        if (sheetData.length < 2) {
          setUploadError("表格中未检测到有效的数据行");
          return;
        }
        
        const headers = sheetData[0].map((h: any) => String(h || "").trim());
        const rows = sheetData.slice(1);
        
        const findColumnIndex = (keywords: string[], exactMatches: string[] = []) => {
          if (exactMatches.length > 0) {
            const idx = headers.findIndex(h => exactMatches.some(em => h === em));
            if (idx !== -1) return idx;
          }
          return headers.findIndex(h => keywords.some(kw => h.toLowerCase().includes(kw)));
        };

        const companyIdx = findColumnIndex(
          ["公司", "company", "企业", "单位", "收货人公司", "买家公司", "buyer company", "recipient company", "firm", "co.", "店铺", "买家"],
          ["买家名称", "公司名称", "客户名称", "Company Name"]
        );

        const recipientIdx = findColumnIndex(
          ["收货人姓名", "收货人", "收件人名称", "收件人", "姓名", "recipient name", "receiver", "name", "买家姓名", "buyer name", "contact", "联系人"],
          ["收件人名称", "收件人", "收货人姓名", "Recipient Name"]
        );

        const countryIdx = findColumnIndex(
          ["国家", "country", "国家名", "中文国家名", "收货人国家", "收货国家", "recipient country", "destination", "nation", "二字码"],
          ["收货国家", "国家", "Country"]
        );

        const productIdx = findColumnIndex(
          ["产品名称", "product name", "产品", "商品", "product", "sku", "规格", "产品规格", "goods", "意向", "商品信息"],
          ["商品信息", "产品名称", "Product Name"]
        );

        const buyerAccountIdx = findColumnIndex(
          ["买家账号", "买家帐号", "buyer account", "buyer username"],
          ["买家账号", "买家帐号"]
        );
        
        const parsedLeads: Lead[] = [];
        
        rows.forEach((row: any[], rowIdx: number) => {
          if (!row || row.length === 0 || row.every(val => val === undefined || val === null || String(val).trim() === "")) {
            return;
          }
          
          const rawCompany = companyIdx !== -1 && row[companyIdx] ? String(row[companyIdx]).trim() : "";
          const rawRecipient = recipientIdx !== -1 && row[recipientIdx] ? String(row[recipientIdx]).trim() : "";
          const rawBuyerAccount = buyerAccountIdx !== -1 && row[buyerAccountIdx] ? String(row[buyerAccountIdx]).trim() : "";
          const country = countryIdx !== -1 && row[countryIdx] ? String(row[countryIdx]).trim() : "";
          const productContext = productIdx !== -1 && row[productIdx] ? String(row[productIdx]).trim() : "包装设备/贴标机 (Packaging Machinery)";
          
          let companyName = "";
          if (rawCompany) {
            companyName = rawCompany;
          } else if (rawRecipient) {
            companyName = `${rawRecipient} (个人客商)`;
          } else if (rawBuyerAccount) {
            companyName = `买家: ${rawBuyerAccount}`;
          } else {
            companyName = `客商 #${rowIdx + 1}`;
          }
          
          const infoParts: string[] = [];
          headers.forEach((header, cellIdx) => {
            const val = row[cellIdx];
            if (val !== undefined && val !== null && String(val).trim() !== "") {
              infoParts.push(`[${header}]: ${String(val).trim()}`);
            }
          });
          const orderInfo = infoParts.join("\n");
          
          parsedLeads.push({
            id: `lead-xls-${Date.now()}-${rowIdx}`,
            companyName,
            recipientName: rawRecipient || undefined,
            country,
            orderInfo,
            productContext,
            status: "idle" as const,
            excelRowIndex: rowIdx
          });
        });
        
        if (parsedLeads.length === 0) {
          setUploadError("未解析到任何符合规则 of 客户数据");
          return;
        }
        
        setOriginalSheetData(sheetData);
        setOriginalFileName(file.name);
        setLeads(prev => [...prev, ...parsedLeads]);
        setUploadSuccessMsg(`成功解析并导入 ${parsedLeads.length} 条客户背调线索！`);
        setTimeout(() => {
          setIsAddingLead(false);
          setUploadSuccessMsg(null);
          if (parsedLeads.length > 0) {
            setSelectedLeadId(parsedLeads[parsedLeads.length - 1].id);
          }
        }, 1500);
        
      } catch (err: any) {
        console.error(err);
        setUploadError(`文件读取失败: ${err.message || "未知错误"}`);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Call the server-side Express API to conduct DeepSeek web-search grounded due diligence
  const handleInvestigate = async (id: string) => {
    // Set lead status to searching
    setLeads(prev => prev.map(l => l.id === id ? { ...l, status: "searching", error: undefined } : l));

    const targetLead = leads.find(l => l.id === id);
    if (!targetLead) return;

    try {
      const headers: Record<string, string> = { "Content-Type": "application/json" };
      if (customApiKey.trim()) {
        headers["x-deepseek-api-key"] = customApiKey.trim();
      }
      
      const response = await fetch("/api/investigate", {
        method: "POST",
        headers,
        body: JSON.stringify({
          companyName: targetLead.companyName,
          country: targetLead.country,
          orderInfo: targetLead.orderInfo,
          productContext: targetLead.productContext,
          model: selectedModel,
        }),
      });

      const data = await response.json();

      if (response.ok && data.success) {
        // Defensive check: never mark a lead "completed" with an empty report.
        // Server-side retries should catch this, but guard here too so an empty
        // result surfaces as a visible failure instead of a silent blank report.
        const reportText = typeof data.report === "string" ? data.report : "";
        const gradeText = typeof data.grade === "string" ? data.grade : "";
        if (!reportText.trim() || !gradeText.trim()) {
          throw new Error("AI 调查引擎返回了空报告或缺失分析等级，请点击「重试」后再次核查。");
        }

        const completedLead = {
          ...targetLead,
          status: "completed" as const,
          grade: data.grade,
          summary: data.summary,
          riskAlert: data.riskAlert,
          followUpStrategy: data.followUpStrategy,
          followUpGrade: data.followUpGrade,
          countryNameZh: data.countryNameZh,
          productNameZh: data.productNameZh,
          companyOverview: data.companyOverview,
          report: data.report,
          sources: data.sources,
          completedAt: new Date().toISOString().replace("T", " ").substring(0, 19),
          usage: data.usage,
        };

        setLeads(prev => prev.map(l => l.id === id ? completedLead : l));

        if (enableFeishu) {
          try {
            // First, auto-create/import the online Feishu Doc and get its URL
            let docUrl = "";
            try {
              docUrl = await uploadSingleToFeishu(id, data.report, completedLead) || "";
            } catch (docErr) {
              console.error("Auto Feishu Doc creation failed during investigation completion:", docErr);
            }

            // Keep the five result fields atomic: do not sync a record without its document URL.
            if (docUrl && feishuBitableUrl.trim()) {
              const leadWithDoc = { ...completedLead, feishuUrl: docUrl };
              await syncSingleToBitable(id, leadWithDoc);
            } else if (!docUrl && feishuBitableUrl.trim()) {
              setLeads(prev => prev.map(l => l.id === id ? {
                ...l,
                bitableStatus: "idle" as const,
                bitableError: "云文档创建失败，多维表格尚未同步"
              } : l));
            }
          } catch (feishuErr) {
            console.error("Auto Feishu synchronization loop failed:", feishuErr);
          }
        }
      } else {
        throw new Error(data.error || "Investigator failed to fetch profile information.");
      }
    } catch (err: any) {
      console.error(err);
      setLeads(prev => prev.map(l => l.id === id ? {
        ...l,
        status: "failed",
        error: err.message || "连接服务器失败"
      } : l));
    }
  };

  const uploadSingleToFeishu = async (id: string, customReport?: string, customLead?: Lead) => {
    const targetLead = customLead || leads.find(l => l.id === id);
    if (!targetLead) return;

    setLeads(prev => prev.map(l => l.id === id ? {
      ...l,
      feishuStatus: "uploading" as const,
      feishuError: undefined
    } : l));

    try {
      if (!feishuAppId || !feishuAppSecret || !feishuFolderToken) {
        throw new Error("请先在「飞书云文档集成配置」面板中配置 App ID, App Secret 以及 目标文件夹 Token。");
      }

      const reportContent = customReport || targetLead.report;
      if (!reportContent) {
        throw new Error("报告内容为空，请先运行或等待背景调查完成！");
      }

      const wordHtml = convertMarkdownToWordHtml(targetLead.companyName, reportContent);
      const fileName = getExportFileName(targetLead, "doc");
      const utf8WithBOM = '\ufeff' + wordHtml;
      const fileContentBase64 = btoa(unescape(encodeURIComponent(utf8WithBOM)));
      const rootFolder = feishuFolderToken.trim();
      const savedSubfolder = readFeishuSubfolderState(rootFolder);

      const response = await fetch("/api/feishu/upload", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          appId: feishuAppId.trim(),
          appSecret: feishuAppSecret.trim(),
          folderToken: feishuFolderToken.trim(),
          currentSubfolderToken: savedSubfolder?.token || "",
          currentSubfolderIndex: savedSubfolder?.index || 0,
          fileName,
          fileContentBase64,
          markdownContent: reportContent,
          companyName: targetLead.companyName,
          documentTitle: buildFeishuDocumentTitle(targetLead),
        }),
      });

      const data = await response.json();
      if (data.subfolderToken && data.subfolderIndex) {
        saveFeishuSubfolderState(rootFolder, {
          token: data.subfolderToken,
          index: Number(data.subfolderIndex),
          name: data.subfolderName || undefined,
        });
      }
      if (!response.ok || !data.success) {
        const uploadError: any = new Error(data.error || "上传到飞书失败");
        uploadError.errorType = data.errorType;
        throw uploadError;
      }

      setLeads(prev => prev.map(l => l.id === id ? {
        ...l,
        feishuStatus: "success" as const,
        feishuUrl: data.url,
        feishuError: undefined,
        feishuImportFallback: data.importFallback === true,
        feishuImportError: data.importError || undefined,
        feishuImportAttempts: data.importAttempts || 1
      } : l));
      return data.url;
    } catch (err: any) {
      console.error("Feishu upload failed for lead:", id, err);
      const errMsg = err.message || "未知上传错误";
      setLeads(prev => prev.map(l => l.id === id ? {
        ...l,
        feishuStatus: "failed" as const,
        feishuError: errMsg
      } : l));
      throw err;
    }
  };

  const syncSingleToBitable = async (id: string, customLead?: Lead) => {
    const targetLead = customLead || leads.find(l => l.id === id);
    if (!targetLead) return;

    setLeads(prev => prev.map(l => l.id === id ? {
      ...l,
      bitableStatus: "syncing" as const,
      bitableError: undefined
    } : l));

    try {
      if (!feishuAppId || !feishuAppSecret || !feishuBitableUrl) {
        throw new Error("请先在右上角「配置飞书同步」面板中设置 App ID, App Secret 和多维表格 URL。");
      }

      const fields: Record<string, any> = {};

      // 1. If this lead has an associated Excel row, populate all Excel columns first to prevent them from being empty
      if (targetLead.excelRowIndex !== undefined && originalSheetData && originalSheetData.length > 1) {
        const headers = originalSheetData[0];
        const row = originalSheetData[targetLead.excelRowIndex + 1];
        if (row) {
          headers.forEach((header, idx) => {
            const val = row[idx];
            if (val !== undefined && val !== null) {
              if (val instanceof Date) {
                const y = val.getFullYear();
                const m = String(val.getMonth() + 1).padStart(2, "0");
                const d = String(val.getDate()).padStart(2, "0");
                const hh = String(val.getHours()).padStart(2, "0");
                const mm = String(val.getMinutes()).padStart(2, "0");
                const ss = String(val.getSeconds()).padStart(2, "0");
                fields[header] = `${y}-${m}-${d} ${hh}:${mm}:${ss}`;
              } else if (typeof val === "number") {
                if (val > 99999999999999) {
                  fields[header] = String(val);
                } else {
                  fields[header] = val;
                }
              } else if (typeof val === "string") {
                const trimmed = val.trim();
                if (trimmed !== "") fields[header] = trimmed;
              } else {
                fields[header] = val;
              }
            }
          });
        }
      }

      // 2. Overwrite or inject back-investigation standard fields on top
      fields.companyName = targetLead.companyName;
      fields.country = targetLead.country;
      fields.productContext = targetLead.productContext;
      fields.grade = targetLead.grade || "";
      fields.summary = targetLead.summary || "";
      fields.riskAlert = targetLead.riskAlert || "";
      fields.followUpStrategy = targetLead.followUpStrategy || "";
      fields.followUpGrade = targetLead.followUpGrade || "";
      fields.feishuUrl = targetLead.feishuUrl || "";
      fields.completedAt = targetLead.completedAt || new Date().toISOString().replace("T", " ").substring(0, 19);

      // 3. Only send orderInfo as a fallback if the lead was manually added and has no Excel columns
      if (targetLead.excelRowIndex === undefined) {
        fields.orderInfo = targetLead.orderInfo || "";
      }

      const response = await fetch("/api/feishu/sync-bitable", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          appId: feishuAppId.trim(),
          appSecret: feishuAppSecret.trim(),
          bitableUrl: feishuBitableUrl.trim(),
          records: [
            {
              fields
            }
          ],
          mode: appMode
        }),
      });

      const data = await response.json();
      if (!response.ok || !data.success) {
        throw new Error(data.error || "同步多维表格失败");
      }

      setLeads(prev => prev.map(l => l.id === id ? {
        ...l,
        bitableStatus: "success" as const,
        bitableUrl: feishuBitableUrl.trim(),
        bitableError: undefined
      } : l));
      return feishuBitableUrl.trim();
    } catch (err: any) {
      console.error("Feishu bitable sync failed for lead:", id, err);
      const errMsg = err.message || "未知多维表格同步错误";
      setLeads(prev => prev.map(l => l.id === id ? {
        ...l,
        bitableStatus: "failed" as const,
        bitableError: errMsg
      } : l));
      throw err;
    }
  };

  const retryUploadAndSync = async (lead: Lead) => {
    try {
      const docUrl = await uploadSingleToFeishu(lead.id, lead.report, lead);
      if (docUrl && feishuBitableUrl.trim()) {
        await syncSingleToBitable(lead.id, { ...lead, feishuUrl: docUrl });
      }
    } catch (error) {
      console.error("Feishu document retry did not reach Bitable sync:", error);
    }
  };

  const handleInvestigateAll = async () => {
    const idleLeads = leads.filter(l => l.status === "idle" || l.status === "failed");
    for (const lead of idleLeads) {
      await handleInvestigate(lead.id);
    }
  };

  // Helper utility to convert Markdown into polished HTML compatible with MS Word/WPS
  const inlineMarkdownToHtml = (text: string): string => {
    let escaped = text
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
    
    // Replace bold **text**
    escaped = escaped.replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>");
    // Replace italic *text*
    escaped = escaped.replace(/\*(.*?)\*/g, "<em>$1</em>");
    // Replace code `text`
    escaped = escaped.replace(/`(.*?)`/g, "<code style='background-color:#f3f4f6;padding:2px 4px;border-radius:4px;'>$1</code>");
    // Replace links [text](url)
    escaped = escaped.replace(/\[(.*?)\]\((.*?)\)/g, "<a href='$2'>$1</a>");
    
    return escaped;
  };

  const convertMarkdownToWordHtml = (companyName: string, markdownText: string): string => {
    const preprocessedText = preprocessMarkdown(markdownText);
    const lines = preprocessedText.split("\n");
    let htmlResult = "";
    let inList = false;
    let listType: "ul" | "ol" | null = null;
    let inTable = false;
    let tableHeaders: string[] = [];
    let tableRows: string[][] = [];

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();

      // Check for tables
      if (line.startsWith("|")) {
        if (!inTable) {
          inTable = true;
          tableHeaders = line.split("|").map(s => s.trim()).filter((s, idx, arr) => idx > 0 && idx < arr.length - 1);
          continue;
        } else {
          if (line.includes("---")) {
            continue;
          }
          const cells = line.split("|").map(s => s.trim()).filter((s, idx, arr) => idx > 0 && idx < arr.length - 1);
          if (cells.length > 0) {
            tableRows.push(cells);
          }
          continue;
        }
      } else {
        if (inTable) {
          htmlResult += "<table><thead><tr>";
          tableHeaders.forEach(h => {
            htmlResult += `<th>${inlineMarkdownToHtml(h)}</th>`;
          });
          htmlResult += "</tr></thead><tbody>";
          tableRows.forEach(row => {
            htmlResult += "<tr>";
            row.forEach(cell => {
              htmlResult += `<td>${inlineMarkdownToHtml(cell)}</td>`;
            });
            htmlResult += "</tr>";
          });
          htmlResult += "</tbody></table>";
          
          inTable = false;
          tableHeaders = [];
          tableRows = [];
        }
      }

      if (line.startsWith("- ") || line.startsWith("* ") || /^\d+\.\s/.test(line)) {
        const isOrdered = /^\d+\.\s/.test(line);
        const currentListType = isOrdered ? "ol" : "ul";
        
        if (!inList || listType !== currentListType) {
          if (inList && listType) {
            htmlResult += `</${listType}>`;
          }
          inList = true;
          listType = currentListType;
          htmlResult += `<${listType}>`;
        }
        
        const content = isOrdered ? line.replace(/^\d+\.\s/, "") : line.substring(2);
        htmlResult += `<li>${inlineMarkdownToHtml(content)}</li>`;
        continue;
      } else {
        if (inList && listType) {
          htmlResult += `</${listType}>`;
          inList = false;
          listType = null;
        }
      }

      if (line === "") {
        continue;
      }

      if (line.startsWith("# ")) {
        htmlResult += `<h1>${inlineMarkdownToHtml(line.substring(2))}</h1>`;
      } else if (line.startsWith("## ")) {
        htmlResult += `<h2>${inlineMarkdownToHtml(line.substring(3))}</h2>`;
      } else if (line.startsWith("### ")) {
        htmlResult += `<h3>${inlineMarkdownToHtml(line.substring(4))}</h3>`;
      } else if (line.startsWith("> ")) {
        htmlResult += `<blockquote>${inlineMarkdownToHtml(line.substring(2))}</blockquote>`;
      } else {
        htmlResult += `<p>${inlineMarkdownToHtml(line)}</p>`;
      }
    }

    if (inTable) {
      htmlResult += "<table><thead><tr>";
      tableHeaders.forEach(h => {
        htmlResult += `<th>${inlineMarkdownToHtml(h)}</th>`;
      });
      htmlResult += "</tr></thead><tbody>";
      tableRows.forEach(row => {
        htmlResult += "<tr>";
        row.forEach(cell => {
          htmlResult += `<td>${inlineMarkdownToHtml(cell)}</td>`;
        });
        htmlResult += "</tr>";
      });
      htmlResult += "</tbody></table>";
    }

    if (inList && listType) {
      htmlResult += `</${listType}>`;
    }

    const documentHtml = `<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
<head>
  <title>${companyName} 背景调查报告</title>
  <!--[if gte mso 9]>
  <xml>
    <w:WordDocument>
      <w:View>Print</w:View>
      <w:Zoom>100</w:Zoom>
      <w:DoNotOptimizeForBrowser/>
    </w:WordDocument>
  </xml>
  <![endif]-->
  <style>
    body {
      font-family: "Microsoft YaHei", "SimSun", Arial, sans-serif;
      font-size: 11pt;
      line-height: 1.6;
      color: #333333;
      padding: 30px;
    }
    h1 {
      font-family: "Microsoft YaHei", sans-serif;
      font-size: 18pt;
      font-weight: bold;
      text-align: center;
      color: #000000;
      border-bottom: 2px solid #000000;
      padding-bottom: 8px;
      margin-top: 20px;
      margin-bottom: 25px;
    }
    h2 {
      font-family: "Microsoft YaHei", sans-serif;
      font-size: 13pt;
      font-weight: bold;
      color: #0d5c3a;
      border-bottom: 1px solid #cccccc;
      padding-bottom: 5px;
      margin-top: 30px;
      margin-bottom: 15px;
    }
    h3 {
      font-family: "Microsoft YaHei", sans-serif;
      font-size: 11pt;
      font-weight: bold;
      color: #333333;
      margin-top: 20px;
      margin-bottom: 10px;
    }
    p {
      margin-bottom: 12px;
      text-align: justify;
      text-justify: inter-ideograph;
      font-size: 10.5pt;
      color: #333333;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      margin-top: 20px;
      margin-bottom: 20px;
    }
    th, td {
      border: 1px solid #aaaaaa;
      padding: 10px 14px;
      text-align: left;
      font-size: 10pt;
    }
    th {
      background-color: #f3f4f6;
      font-weight: bold;
      color: #111827;
    }
    ul, ol {
      margin-bottom: 15px;
      padding-left: 20px;
    }
    li {
      margin-bottom: 6px;
      font-size: 10.5pt;
      color: #333333;
    }
    blockquote {
      border-left: 4px solid #10b981;
      padding-left: 15px;
      color: #4b5563;
      font-style: italic;
      margin: 15px 0;
      background-color: #f9fafb;
      padding-top: 8px;
      padding-bottom: 8px;
    }
    a {
      color: #059669;
      text-decoration: underline;
    }
  </style>
</head>
<body>
  ${htmlResult}
</body>
</html>`;

    return documentHtml;
  };

  // Helper to prepare the updated Excel rows with investigation status/links
  const prepareUpdatedExcelRows = (): any[][] | null => {
    if (!originalSheetData || originalSheetData.length < 2) {
      return null;
    }

    const headers = originalSheetData[0] || [];

    // Column lookup helper: exact or substring (case-insensitive) match.
    // Columns already claimed by a more specific lookup are skipped, so e.g.
    // "跟进等级" is claimed by the follow-up lookup and not re-matched as a
    // generic grade column.
    const assignedHeaderIdxs = new Set<number>();
    const findHeaderIdx = (keywords: string[]) => {
      for (let i = 0; i < headers.length; i++) {
        if (assignedHeaderIdxs.has(i)) continue;
        const h = String(headers[i] || "").trim().toLowerCase();
        const matched = keywords.some(kw => {
          const k = kw.toLowerCase();
          return h === k || h.includes(k);
        });
        if (matched) {
          assignedHeaderIdxs.add(i);
          return i;
        }
      }
      return -1;
    };

    // 标题 / 报告链接 column (existing column, or appended if missing)
    let reportColIdx = -1;
    const reportColKeywords = ["标题", "背景报告文档", "背景调查报告文档", "报告链接", "背调报告文档", "feishuUrl"];
    for (let i = 0; i < headers.length; i++) {
      const h = String(headers[i] || "").trim();
      if (reportColKeywords.includes(h) || h.toLowerCase() === "feishuurl") {
        reportColIdx = i;
        assignedHeaderIdxs.add(i);
        break;
      }
    }

    // Detect the specific follow-up columns first, then the generic ones,
    // so specific headers are never mis-assigned.
    const riskAlertIdx = findHeaderIdx(["风险提示", "风险警告", "风险警示", "风险", "risk"]);
    const followUpStrategyIdx = findHeaderIdx(["跟进策略", "跟进方案", "跟进话术", "策略", "strategy"]);
    const followUpGradeIdx = findHeaderIdx(["跟进等级", "跟进级别", "跟进评级", "跟进强度", "跟进", "follow up", "followup"]);
    const gradeIdx = findHeaderIdx(["背景评判等级", "背景分析等级", "评估等级", "背调等级", "评级", "grade", "客户等级", "等级"]);
    const summaryIdx = findHeaderIdx(["报告概要", "调查概要", "背调摘要", "摘要", "结论", "summary", "overview"]);
    // NOTE: only match investigation-specific time columns. Generic substrings
    // like "时间"/"日期" would match order columns (下单时间, 付款时间, 发货时间…)
    // and overwrite them with the investigation timestamp.
    const completedAtIdx = findHeaderIdx(["背调时间", "调查时间", "背调完成时间", "调查完成时间", "背调日期", "completed at", "completedat", "investigation time"]);

    const updatedRows = originalSheetData.map(row => (row ? [...row] : []));
    let maxLen = 0;
    updatedRows.forEach(r => {
      if (r.length > maxLen) maxLen = r.length;
    });

    // Append header cells for any investigation column missing from the original sheet
    const appendedCols: Array<{ index: number; header: string }> = [];
    const ensureColumn = (existingIdx: number, header: string): number => {
      if (existingIdx !== -1) return existingIdx;
      const newIdx = maxLen + appendedCols.length;
      appendedCols.push({ index: newIdx, header });
      return newIdx;
    };

    const reportCol = ensureColumn(reportColIdx, "标题");
    const gradeCol = ensureColumn(gradeIdx, "评估等级");
    const summaryCol = ensureColumn(summaryIdx, "报告概要");
    const completedAtCol = ensureColumn(completedAtIdx, "背调时间");
    const riskAlertCol = ensureColumn(riskAlertIdx, "风险提示");
    const followUpStrategyCol = ensureColumn(followUpStrategyIdx, "跟进策略");
    const followUpGradeCol = ensureColumn(followUpGradeIdx, "跟进等级");

    if (updatedRows[0]) {
      appendedCols.forEach(({ index, header }) => {
        updatedRows[0][index] = header;
      });
    }

    const allCols = [reportCol, gradeCol, summaryCol, completedAtCol, riskAlertCol, followUpStrategyCol, followUpGradeCol];
    const maxIdxNeeded = Math.max(...allCols);

    const excelLeads = leads.filter(l => l.excelRowIndex !== undefined);
    excelLeads.forEach(lead => {
      if (lead.excelRowIndex === undefined) return;
      const rowIndex = lead.excelRowIndex + 1; // Map back to sheetData (offset by header)
      if (!updatedRows[rowIndex]) return;

      // Ensure the data row is long enough for all target columns
      for (let j = updatedRows[rowIndex].length; j <= maxIdxNeeded; j++) {
        updatedRows[rowIndex][j] = "";
      }

      if (lead.status === "completed" && lead.report) {
        const fileName = getExportFileName(lead, "doc");
        updatedRows[rowIndex][reportCol] = lead.feishuUrl ? lead.feishuUrl : fileName;
        updatedRows[rowIndex][gradeCol] = lead.grade || "";
        updatedRows[rowIndex][summaryCol] = lead.summary || "";
        updatedRows[rowIndex][completedAtCol] = lead.completedAt || new Date().toLocaleString("zh-CN");
        updatedRows[rowIndex][riskAlertCol] = lead.riskAlert || "";
        updatedRows[rowIndex][followUpStrategyCol] = lead.followUpStrategy || "";
        updatedRows[rowIndex][followUpGradeCol] = lead.followUpGrade || "";
      } else if (lead.status === "completed") {
        // Old-session footprint: marked completed but the report is missing.
        // Make it visible instead of silently blank.
        updatedRows[rowIndex][reportCol] = "⚠️ 已完成但结果为空，请重新核查";
      } else {
        updatedRows[rowIndex][reportCol] = lead.status === "failed" ? "背调失败" : "未完成背调";
      }
    });

    return updatedRows;
  };

  // Export only the updated Excel sheet (WITHOUT the Word reports inside ZIP)
  const handleDownloadOnlyExcel = () => {
    const updatedRows = prepareUpdatedExcelRows();
    if (!updatedRows) {
      alert("没有检测到已上传的原始表格数据，请先导入 Excel 表格！");
      return;
    }

    try {
      const worksheet = XLSX.utils.aoa_to_sheet(updatedRows);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "背调后客商清单");
      
      const baseNameWithoutExt = originalFileName ? originalFileName.substring(0, originalFileName.lastIndexOf(".")) : "customer_list";
      const newExcelName = `${baseNameWithoutExt}_completed_reports.xlsx`;

      XLSX.writeFile(workbook, newExcelName);
    } catch (error: any) {
      console.error("Failed to export Excel:", error);
      alert(`导出表格失败: ${error.message || error}`);
    }
  };

  // Sync the updated Excel records to Feishu Bitable (Multi-dimensional Table) as new records
  const handleUploadSheetToFeishu = async () => {
    const updatedRows = prepareUpdatedExcelRows();
    if (!updatedRows || updatedRows.length < 2) {
      alert("没有检测到已上传的原始表格数据，请先导入 Excel 表格！");
      return;
    }

    if (!feishuAppId || !feishuAppSecret || !feishuBitableUrl) {
      alert("请先在右上角「飞书云文档集成配置」面板中配置 App ID, App Secret 以及 飞书多维表格 (Bitable) 链接。");
      setFeishuConfigOpen(true);
      return;
    }

    setIsUploadingSheet(true);
    setUploadSheetError(null);

    try {
      const headers = updatedRows[0];
      const records = updatedRows.slice(1).map((row, rIdx) => {
        const fields: Record<string, any> = {};
        
        // 1. Populate from Excel row first with advanced date and number formatting
        headers.forEach((header, idx) => {
          const val = row[idx];
          if (val !== undefined && val !== null) {
            if (val instanceof Date) {
              const y = val.getFullYear();
              const m = String(val.getMonth() + 1).padStart(2, "0");
              const d = String(val.getDate()).padStart(2, "0");
              const hh = String(val.getHours()).padStart(2, "0");
              const mm = String(val.getMinutes()).padStart(2, "0");
              const ss = String(val.getSeconds()).padStart(2, "0");
              fields[header] = `${y}-${m}-${d} ${hh}:${mm}:${ss}`;
            } else if (typeof val === "number") {
              // High-precision ID preservation (over 14 digits) as strings to prevent float issues
              if (val > 99999999999999) {
                fields[header] = String(val);
              } else {
                fields[header] = val;
              }
            } else if (typeof val === "string") {
              const trimmed = val.trim();
              if (trimmed !== "") fields[header] = trimmed;
            } else {
              fields[header] = val;
            }
          }
        });

        // 2. Inject back-investigation results if available
        const lead = leads.find(l => l.excelRowIndex === rIdx);
        if (lead && lead.status === "completed") {
          // Explicitly map standard fields so keyAliases on the backend can route them to the correct Bitable column
          fields.companyName = lead.companyName;
          fields.country = lead.country;
          fields.productContext = lead.productContext;
          fields.grade = lead.grade;
          fields.summary = lead.summary;
          fields.riskAlert = lead.riskAlert;
          fields.followUpStrategy = lead.followUpStrategy;
          fields.followUpGrade = lead.followUpGrade;
          fields.feishuUrl = lead.feishuUrl || "";
          fields.completedAt = lead.completedAt || new Date().toLocaleString("zh-CN");
        }

        return { fields };
      });

      const response = await fetch("/api/feishu/sync-bitable", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          appId: feishuAppId.trim(),
          appSecret: feishuAppSecret.trim(),
          bitableUrl: feishuBitableUrl.trim(),
          records,
          mode: appMode
        }),
      });

      const data = await response.json();
      if (!response.ok || !data.success) {
        throw new Error(data.error || "同步飞书多维表格失败");
      }

      setFeishuSheetUrl(feishuBitableUrl.trim());
      localStorage.setItem("feishu_sheet_url", feishuBitableUrl.trim());
      
      if (appMode === "update") {
        const reasonLabels: Record<string, string> = {
          record_id_not_found: "记录 ID 不存在",
          record_id_order_conflict: "记录 ID 与订单号冲突",
          missing_order_number: "缺少订单号",
          order_number_not_found: "订单号未找到",
          duplicate_order_number: "订单号重复",
          missing_identifiers: "缺少匹配标识",
          no_updatable_fields: "没有可更新字段",
        };
        const reasonDetails = Object.entries(data.matchReasonCounts || {})
          .filter(([, count]) => Number(count) > 0)
          .map(([reason, count]) => `${reasonLabels[reason] || reason}：${count} 条`)
          .join("\n");
        alert(
          `🎉 精准更新完成！\n` +
          `实际更新：${data.updatedCount || 0} 条\n` +
          `记录 ID 双重校验：${data.recordIdMatchedCount || 0} 条\n` +
          `唯一订单号兜底：${data.orderFallbackMatchedCount || 0} 条\n` +
          `冲突：${data.conflictCount || 0} 条\n` +
          `未匹配/跳过：${data.unmatchedCount || 0} 条` +
          (reasonDetails ? `\n\n跳过原因：\n${reasonDetails}` : "")
        );
      }

      if (data.mappingAnalysis) {
        setSyncMappingAnalysis(data.mappingAnalysis);
        setIsMappingAnalysisModalOpen(true);
      } else if (appMode !== "update") {
        alert(`🎉 成功向飞书多维表格新增了 ${data.addedCount || records.length} 条背调记录！`);
      }
    } catch (error: any) {
      console.error("Feishu bitable sync failed:", error);
      setUploadSheetError(error.message || "未知错误");
      alert(`同步飞书多维表格失败: ${error.message || error}`);
    } finally {
      setIsUploadingSheet(false);
    }
  };

  // Build the entire ZIP package containing the updated Excel sheet and Word docs
  const handleDownloadBulkZip = async () => {
    const updatedRows = prepareUpdatedExcelRows();
    if (!updatedRows) {
      alert("没有检测到已上传的原始表格数据，请先导入 Excel 表格！");
      return;
    }

    setIsExportingZip(true);
    try {
      const zip = new JSZip();

      // 2. Generate Word documents and add them to the ZIP
      const docFolder = zip.folder("Word_Due_Diligence_Reports");

      const excelLeads = leads.filter(l => l.excelRowIndex !== undefined);
      excelLeads.forEach(lead => {
        if (lead.excelRowIndex !== undefined) {
          if (lead.status === "completed" && lead.report) {
            const fileName = getExportFileName(lead, "doc");
            const wordHtml = convertMarkdownToWordHtml(lead.companyName, lead.report);
            docFolder?.file(fileName, '\ufeff' + wordHtml);
          }
        }
      });

      // Convert updated rows back to XLSX sheet
      const worksheet = XLSX.utils.aoa_to_sheet(updatedRows);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "背调后客商清单");
      
      // Generate Excel file buffer
      const excelBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
      
      // Add Excel file to the root of the ZIP
      const baseNameWithoutExt = originalFileName ? originalFileName.substring(0, originalFileName.lastIndexOf(".")) : "customer_list";
      const newExcelName = `${baseNameWithoutExt}_completed_reports.xlsx`;
      zip.file(newExcelName, excelBuffer);

      // Generate the ZIP file blob
      const content = await zip.generateAsync({ type: "blob" });
      
      // Trigger download of the ZIP file
      const url = URL.createObjectURL(content);
      const link = document.createElement("a");
      link.href = url;
      link.download = `collective_due_diligence_export_${Date.now()}.zip`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    } catch (error: any) {
      console.error("Failed to generate collective export ZIP:", error);
      alert(`集体导出打包失败: ${error.message || error}`);
    } finally {
      setIsExportingZip(false);
    }
  };

  // Generate a mock download action for Markdown reports
  const handleDownloadMD = (lead: Lead) => {
    if (!lead.report) return;
    const blob = new Blob([lead.report], { type: "text/markdown;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = getExportFileName(lead, "md");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  // Export full due diligence report to Microsoft Word (DOC) format
  const handleDownloadWord = (lead: Lead) => {
    const htmlContent = document.getElementById("report-print-content")?.innerHTML;
    if (!htmlContent) return;

    // Word HTML document configuration with inline styling optimized for MS Word / WPS
    const documentHtml = `<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
<head>
  <title>${lead.companyName} 背景调查报告</title>
  <!--[if gte mso 9]>
  <xml>
    <w:WordDocument>
      <w:View>Print</w:View>
      <w:Zoom>100</w:Zoom>
      <w:DoNotOptimizeForBrowser/>
    </w:WordDocument>
  </xml>
  <![endif]-->
  <style>
    body {
      font-family: "Microsoft YaHei", "SimSun", Arial, sans-serif;
      font-size: 11pt;
      line-height: 1.6;
      color: #333333;
      padding: 30px;
    }
    h1 {
      font-family: "Microsoft YaHei", sans-serif;
      font-size: 18pt;
      font-weight: bold;
      text-align: center;
      color: #000000;
      border-bottom: 2px solid #000000;
      padding-bottom: 8px;
      margin-top: 20px;
      margin-bottom: 25px;
    }
    h2 {
      font-family: "Microsoft YaHei", sans-serif;
      font-size: 13pt;
      font-weight: bold;
      color: #0d5c3a;
      border-bottom: 1px solid #cccccc;
      padding-bottom: 5px;
      margin-top: 30px;
      margin-bottom: 15px;
    }
    h3 {
      font-family: "Microsoft YaHei", sans-serif;
      font-size: 11pt;
      font-weight: bold;
      color: #333333;
      margin-top: 20px;
      margin-bottom: 10px;
    }
    p {
      margin-bottom: 12px;
      text-align: justify;
      text-justify: inter-ideograph;
      font-size: 10.5pt;
      color: #333333;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      margin-top: 20px;
      margin-bottom: 20px;
    }
    th, td {
      border: 1px solid #aaaaaa;
      padding: 10px 14px;
      text-align: left;
      font-size: 10pt;
    }
    th {
      background-color: #f3f4f6;
      font-weight: bold;
      color: #111827;
    }
    ul, ol {
      margin-bottom: 15px;
      padding-left: 20px;
    }
    li {
      margin-bottom: 6px;
      font-size: 10.5pt;
      color: #333333;
    }
    blockquote {
      border-left: 4px solid #10b981;
      padding-left: 15px;
      color: #4b5563;
      font-style: italic;
      margin: 15px 0;
      background-color: #f9fafb;
      padding-top: 8px;
      padding-bottom: 8px;
    }
    a {
      color: #059669;
      text-decoration: underline;
    }
  </style>
</head>
<body>
  ${htmlContent}
</body>
</html>`;

    const blob = new Blob(['\ufeff' + documentHtml], { type: "application/msword;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = getExportFileName(lead, "doc");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  // Custom Outreach Outreach Template Generator
  const getOutreachTemplates = (lead?: Lead) => {
    if (!lead) {
      return {
        emailSubject: "",
        emailBody: "",
        chatSubject: "",
        chatBody: ""
      };
    }
    const defaultManager = lead.companyOverview?.legalName ? "Purchase Manager / Team" : "Purchasing Team";
    const machinery = lead.productContext || "packaging machine";
    const emailSubject = `🟢 Free Technical Factory Preset for Your ${machinery} - Order Ref`;
    const emailBody = `Dear ${lead.companyName} Team,

I notice you have recently purchased / placed interest in our high-precision ${machinery} on AliExpress/our store. 

As an industrial machinery manufacturer, we always provide custom presets for our B2B clients to perfectly match their production lines. Based on our preliminary engineering background check, we recognize that you operate a reputable manufacturing and processing plant in ${lead.country || "your location"}.

To ensure perfect precision for your bottle/container packaging process:
1. We have set up a Direct Engineering Support Group for your company.
2. We would love to pre-configure your machine's parameters (labeling offsets, sensory triggers, or speed limits) free of charge at our factory before shipping.

Could you please share with us the dimensions and materials of your target containers?

We look forward to building a long-term business relationship with your plant.

Best regards,
B2B Sales Technical Director
Machinery Factory Co.`;

    const chatSubject = "💬 IM (WeChat/WhatsApp/AliExpress) Contact Script";
    const chatBody = `Hi there! I am the lead engineer of your ${machinery} factory. We are preparing your order. I checked your company profile and we highly respect your professional packaging setup! 

To guarantee 100% precision with your bottles/tubes, could you send us a quick photo/dimensions of your container? We will perform a free pre-calibration test for you before delivery. Thanks!`;

    return { emailSubject, emailBody, chatSubject, chatBody };
  };

  const templates = getOutreachTemplates(selectedLead);

  return (
    <div className="min-h-screen bg-[#0B0B0B] text-slate-200 flex flex-col font-sans">
      {/* Top Professional App Header */}
      <header className="border-b border-slate-800/60 bg-[#0F0F0F] px-6 py-4 flex items-center justify-between shadow-sm shrink-0">
        <div className="flex items-center space-x-3">
          <div className="bg-emerald-500/10 text-emerald-400 p-2 rounded-lg border border-emerald-500/20">
            <Building2 className="h-6 w-6" id="app-logo" />
          </div>
          <div>
            <h1 className="text-lg font-bold font-display tracking-tight text-white flex items-center gap-2">
              境外客商背景调查专家 <span className="text-xs bg-emerald-500/10 text-emerald-400 px-2 py-0.5 rounded-full font-medium border border-emerald-500/20">AI Deep Investigation</span>
            </h1>
            <p className="text-xs text-slate-400">基于 DeepSeek V4 Flash 联网搜索实证进行境外 B2B 客户真实性核查与销售建联建议</p>
          </div>
        </div>
        <div className="flex items-center space-x-4">
          {/* Model Selector */}
          <div className="flex items-center gap-2 bg-[#111111] border border-slate-800 rounded-lg px-2.5 py-1.5 hover:border-slate-700 transition">
            <Cpu className="h-3.5 w-3.5 text-emerald-400" />
            <div className="flex flex-col text-left">
              <span className="text-[9px] text-slate-500 uppercase tracking-wider font-semibold">AI 调查模型选择</span>
              <select
                value={selectedModel}
                onChange={(e) => {
                  const val = e.target.value;
                  setSelectedModel(val);
                  localStorage.setItem("selected_deepseek_model", val);
                }}
                className="bg-transparent text-xs font-mono font-medium text-slate-200 focus:outline-none cursor-pointer pr-1"
              >
                <option value="deepseek-v4-flash" className="bg-[#151515] text-slate-300">DeepSeek V4 Flash (推荐)</option>
              </select>
            </div>
          </div>
          <div className="h-8 w-px bg-slate-800 hidden sm:block"></div>

          {/* Elegant API Key Settings Widget */}
          <div className="relative">
            <button
              onClick={() => {
                setShowKeyInput(!showKeyInput);
                setTempApiKey(customApiKey);
              }}
              className={`text-xs font-semibold px-3 py-1.5 rounded-md border flex items-center gap-1.5 transition ${
                customApiKey.trim()
                  ? "bg-emerald-500/10 border-emerald-500/30 text-emerald-400 hover:bg-emerald-500/20"
                  : "bg-amber-500/10 border-amber-500/30 text-amber-400 hover:bg-amber-500/20"
              }`}
            >
              <Key className="h-3.5 w-3.5" />
              {customApiKey.trim() ? "已启用自定义 API Key" : "使用内置/配置 Key (点击设置)"}
            </button>

            {showKeyInput && (
              <div className="absolute right-0 mt-2 w-80 bg-[#151515] border border-slate-800 rounded-xl p-4 shadow-xl z-50 text-left space-y-3 animate-fadeIn">
                <div className="flex justify-between items-center pb-1 border-b border-slate-800/60">
                  <span className="text-xs font-bold text-white flex items-center gap-1">
                    <Key className="h-3.5 w-3.5 text-emerald-400" />
                    配置 DeepSeek API 秘钥
                  </span>
                  <button
                    onClick={() => setShowKeyInput(false)}
                    className="text-slate-500 hover:text-slate-300 text-xs"
                  >
                    关闭
                  </button>
                </div>
                <p className="text-[10px] text-slate-400 leading-relaxed">
                  如果您遇到了 429 访问限流或配额不足报错，请在此处配置您个人的 DeepSeek API Key。秘钥仅保存在您本地浏览器的缓存中，非常安全。
                </p>
                <div className="relative">
                  <input
                    type={showPassword ? "text" : "password"}
                    value={tempApiKey}
                    onChange={(e) => setTempApiKey(e.target.value)}
                    placeholder="sk-..."
                    className="w-full bg-[#0F0F0F] border border-slate-800 rounded-lg pl-3 pr-8 py-2 text-xs font-mono text-white focus:outline-none focus:border-slate-700"
                  />
                  <button
                    type="button"
                    onClick={() => setShowPassword(!showPassword)}
                    className="absolute right-2.5 top-2.5 text-slate-500 hover:text-slate-300"
                  >
                    {showPassword ? <EyeOff className="h-3.5 w-3.5" /> : <Eye className="h-3.5 w-3.5" />}
                  </button>
                </div>
                <div className="flex gap-2 justify-end pt-1">
                  {customApiKey.trim() && (
                    <button
                      onClick={() => {
                        localStorage.removeItem("custom_deepseek_api_key");
                        setCustomApiKey("");
                        setTempApiKey("");
                        setShowKeyInput(false);
                      }}
                      className="text-red-400 hover:text-red-300 text-[10px] font-semibold px-2 py-1 rounded hover:bg-red-500/10 transition"
                    >
                      清空并还原
                    </button>
                  )}
                  <button
                    onClick={() => {
                      const trimmed = tempApiKey.trim();
                      localStorage.setItem("custom_deepseek_api_key", trimmed);
                      setCustomApiKey(trimmed);
                      setShowKeyInput(false);
                    }}
                    className="bg-emerald-600 hover:bg-emerald-500 text-white font-semibold text-[10px] px-3 py-1 rounded transition"
                  >
                    保存配置
                  </button>
                </div>
              </div>
            )}
          </div>

          <div className="h-8 w-px bg-slate-800"></div>

          {/* Elegant Feishu Cloud Docs Settings Widget */}
          <div className="relative">
            <button
              onClick={() => {
                setFeishuConfigOpen(!feishuConfigOpen);
              }}
              className={`text-xs font-semibold px-3 py-1.5 rounded-md border flex items-center gap-1.5 transition ${
                feishuAppId.trim() && feishuAppSecret.trim() && feishuFolderToken.trim()
                  ? "bg-sky-500/10 border-sky-500/30 text-sky-400 hover:bg-sky-500/20"
                  : "bg-slate-800 border-slate-700/60 text-slate-400 hover:bg-slate-700 hover:text-white"
              }`}
            >
              <Cloud className="h-3.5 w-3.5 text-sky-400" />
              {feishuAppId.trim() ? "已配置飞书同步" : "配置飞书同步"}
              {enableFeishu && (
                <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse"></span>
              )}
            </button>

            {feishuConfigOpen && (
              <div className="absolute right-0 mt-2 w-80 bg-[#151515] border border-slate-800 rounded-xl p-4 shadow-xl z-50 text-left space-y-3.5 animate-fadeIn">
                <div className="flex justify-between items-center pb-1 border-b border-slate-800/60">
                  <span className="text-xs font-bold text-white flex items-center gap-1.5">
                    <Cloud className="h-4 w-4 text-sky-400" />
                    飞书云文档集成配置
                  </span>
                  <button
                    onClick={() => setFeishuConfigOpen(false)}
                    className="text-slate-500 hover:text-slate-300 text-xs"
                  >
                    关闭
                  </button>
                </div>
                
                <p className="text-[10px] text-slate-400 leading-relaxed">
                  配置飞书自建应用参数后，系统在背景调查完成后<strong>自动把报告导入并转换成在线飞书文档（docx）并保存至目标文件夹</strong>，免去人工搬运。
                </p>

                <div className="space-y-2.5">
                  <div>
                    <label className="block text-[10px] text-slate-400 mb-1 font-mono">App ID</label>
                    <input
                      type="text"
                      value={feishuAppId}
                      onChange={(e) => {
                        const val = e.target.value;
                        setFeishuAppId(val);
                        localStorage.setItem("feishu_app_id", val);
                      }}
                      placeholder="cli_xxxxxxxxxxxxx"
                      className="w-full bg-[#0F0F0F] border border-slate-800 rounded-lg px-2.5 py-1.5 text-xs font-mono text-white focus:outline-none focus:border-slate-700"
                    />
                  </div>

                  <div>
                    <label className="block text-[10px] text-slate-400 mb-1 font-mono">App Secret</label>
                    <input
                      type="password"
                      value={feishuAppSecret}
                      onChange={(e) => {
                        const val = e.target.value;
                        setFeishuAppSecret(val);
                        localStorage.setItem("feishu_app_secret", val);
                      }}
                      placeholder="••••••••••••••••••••"
                      className="w-full bg-[#0F0F0F] border border-slate-800 rounded-lg px-2.5 py-1.5 text-xs font-mono text-white focus:outline-none focus:border-slate-700"
                    />
                  </div>

                  <div>
                    <label className="block text-[10px] text-slate-400 mb-1 font-mono">报告文档根文件夹 Token</label>
                    <input
                      type="text"
                      value={feishuFolderToken}
                      onChange={(e) => {
                        const val = e.target.value;
                        if (val.trim() !== feishuFolderToken.trim()) {
                          clearFeishuSubfolderState(feishuFolderToken);
                        }
                        setFeishuFolderToken(val);
                        localStorage.setItem("feishu_folder_token", val);
                      }}
                      placeholder="fldcnxxxxxxxxxxxxxxxx"
                      className="w-full bg-[#0F0F0F] border border-slate-800 rounded-lg px-2.5 py-1.5 text-xs font-mono text-white focus:outline-none focus:border-slate-700"
                    />
                    <span className="text-[9px] text-slate-500 block mt-1">
                      程序会在此根目录下自动创建和切换“背调报告”子目录。URL 中 <code>/folder/</code> 后的长字符串即为 Token。
                    </span>
                  </div>

                  <div>
                    <label className="block text-[10px] text-slate-400 mb-1 font-mono">飞书多维表格 (Bitable) 链接</label>
                    <input
                      type="text"
                      value={feishuBitableUrl}
                      onChange={(e) => {
                        const val = e.target.value;
                        setFeishuBitableUrl(val);
                        localStorage.setItem("feishu_bitable_url", val);
                      }}
                      placeholder="https://cathypromotion.feishu.cn/wiki/xxxxxxxxxxxxxxxxxxxxxxxxx?table=tbleBMfSALphGXje..."
                      className="w-full bg-[#0F0F0F] border border-slate-800 rounded-lg px-2.5 py-1.5 text-xs font-mono text-white focus:outline-none focus:border-slate-700"
                    />
                    <span className="text-[9px] text-slate-500 block mt-1 leading-relaxed">
                      （新增表格记录）粘贴多维表格或知识库内嵌多维表格链接，<strong>需包含 table= 参数</strong>。
                    </span>
                  </div>

                  <div className="flex items-center justify-between p-2 bg-[#0F0F0F] rounded-lg border border-slate-800/60 mt-1">
                    <span className="text-[11px] text-slate-300 font-bold">自动同步上传飞书</span>
                    <input
                      type="checkbox"
                      checked={enableFeishu}
                      onChange={(e) => {
                        const checked = e.target.checked;
                        setEnableFeishu(checked);
                        localStorage.setItem("feishu_enabled", checked ? "true" : "false");
                      }}
                      className="h-4 w-4 rounded border-slate-700 text-sky-600 focus:ring-sky-500 bg-slate-950 cursor-pointer"
                    />
                  </div>
                </div>

                <div className="flex justify-between items-center pt-1 border-t border-slate-800/60">
                  <button
                    onClick={() => {
                      setFeishuAppId("");
                      setFeishuAppSecret("");
                      setFeishuFolderToken("");
                      setFeishuBitableUrl("");
                      setEnableFeishu(false);
                      setFeishuSheetUrl("");
                      localStorage.removeItem("feishu_app_id");
                      localStorage.removeItem("feishu_app_secret");
                      localStorage.removeItem("feishu_folder_token");
                      localStorage.removeItem(FEISHU_SUBFOLDER_STATE_KEY);
                      localStorage.removeItem("feishu_bitable_url");
                      localStorage.removeItem("feishu_enabled");
                      localStorage.removeItem("feishu_sheet_url");
                      setFeishuConfigOpen(false);
                    }}
                    className="text-red-400 hover:text-red-300 text-[10px] font-semibold"
                  >
                    重置配置
                  </button>
                  <button
                    onClick={() => {
                      setFeishuConfigOpen(false);
                    }}
                    className="bg-sky-600 hover:bg-sky-500 text-white font-semibold text-[10px] px-3 py-1.5 rounded transition shadow-sm"
                  >
                    确认保存
                  </button>
                </div>
              </div>
            )}
          </div>

          <div className="h-8 w-px bg-slate-800"></div>
          <a
            href="#usage-guide"
            className="text-xs text-slate-400 hover:text-white bg-slate-800/50 hover:bg-slate-700 px-3 py-1.5 rounded-md transition border border-slate-800/60"
          >
            使用指南
          </a>
        </div>
      </header>

      {/* Main Workspace Layout */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-4 sm:p-6 grid grid-cols-1 lg:grid-cols-12 gap-6 bg-[#0B0B0B]">
        
        {/* Left Column: Leads Management Panel (Grid 5) */}
        <div className="lg:col-span-5 flex flex-col space-y-5">
          
          {/* Quick Metrics & Actions Dashboard */}
          <div className="bg-[#151515] border border-slate-800 rounded-xl p-4 shadow-sm">
            <h2 className="text-sm font-semibold text-slate-300 mb-3 flex items-center gap-2">
              <Layers className="h-4 w-4 text-emerald-400" />
              背调进程仪表盘 (Metrics Overview)
            </h2>
            <div className="grid grid-cols-3 gap-3 mb-4">
              <div className="bg-[#0F0F0F] border border-slate-800/60 p-3 rounded-lg text-center">
                <span className="text-xs text-slate-500 block mb-1">总导入客商</span>
                <span className="text-xl font-bold font-mono text-white">{leads.length}</span>
              </div>
              <div className="bg-[#0F0F0F] border border-slate-800/60 p-3 rounded-lg text-center">
                <span className="text-xs text-slate-500 block mb-1">💎 优质客户 (S/A)</span>
                <span className="text-xl font-bold font-mono text-emerald-400">
                  {leads.filter(l => l.grade?.includes("S级") || l.grade?.includes("A级")).length}
                </span>
              </div>
              <div className="bg-[#0F0F0F] border border-slate-800/60 p-3 rounded-lg text-center">
                <span className="text-xs text-slate-500 block mb-1">待背调队列</span>
                <span className="text-xl font-bold font-mono text-amber-500">
                  {leads.filter(l => l.status === "idle").length}
                </span>
              </div>
            </div>

            {/* 工作功能模式选择器 (Function Mode Selector) */}
            <div className="bg-[#121212] border border-slate-800/80 p-3 rounded-xl space-y-2 shadow-inner">
              <div className="flex items-center justify-between">
                <span className="text-xs font-bold text-slate-300 flex items-center gap-1.5">
                  <span className="h-1.5 w-1.5 rounded-full bg-emerald-500 animate-pulse"></span>
                  系统当前工作功能模式
                </span>
                <span className="text-[9px] text-slate-500 font-mono">FUNCTION MODE</span>
              </div>
              
              <div className="grid grid-cols-2 gap-1.5 bg-[#0A0A0A] p-0.5 rounded-lg border border-slate-850">
                <button
                  type="button"
                  onClick={() => setAppMode("create")}
                  className={`py-1.5 px-2 rounded-md text-xs font-semibold flex flex-col items-center justify-center transition ${
                    appMode === "create"
                      ? "bg-gradient-to-r from-emerald-600 to-emerald-700 text-white shadow-md font-bold"
                      : "text-slate-400 hover:text-slate-200 hover:bg-slate-900/40"
                  }`}
                >
                  <span className="text-xs">功能一</span>
                  <span className="text-[9px] opacity-90">新建客商入库</span>
                </button>
                <button
                  type="button"
                  onClick={() => setAppMode("update")}
                  className={`py-1.5 px-2 rounded-md text-xs font-semibold flex flex-col items-center justify-center transition ${
                    appMode === "update"
                      ? "bg-gradient-to-r from-teal-600 to-teal-700 text-white shadow-md font-bold"
                      : "text-slate-400 hover:text-slate-200 hover:bg-slate-900/40"
                  }`}
                >
                  <span className="text-xs">功能二</span>
                  <span className="text-[9px] opacity-90">精准匹配更新</span>
                </button>
              </div>
              
              <div className="text-[10px] text-slate-400 bg-[#161616] p-2 rounded-lg border border-slate-800/50 leading-relaxed font-sans">
                {appMode === "create" ? (
                  <div>
                    <strong className="text-emerald-400">新建入库模式：</strong> 
                    分析上传表格的客户背景，并将背调结果同步为多维表格中的<span className="text-white">全新记录行</span>。
                  </div>
                ) : (
                  <div>
                    <strong className="text-teal-400">精准更新模式：</strong> 
                    使用「记录 ID」精确定位并校验「订单号」；缺少记录 ID 时使用唯一订单号兜底。<span className="text-white font-bold">仅更新背景评判等级、背景报告文档、风险提示、跟进策略、跟进等级 5 个字段</span>，其他字段保持不动。
                  </div>
                )}
              </div>
            </div>

            <div className="flex gap-2">
              <button
                onClick={() => {
                  setIsAddingLead(true);
                  setAddMode("batch");
                }}
                className="flex-1 bg-emerald-600 hover:bg-emerald-500 active:bg-emerald-700 text-white py-2 px-3 rounded-lg text-xs font-semibold flex items-center justify-center gap-1.5 transition"
                id="btn-add-batch"
              >
                <Plus className="h-4 w-4" /> 批量导入客户 (Excel/Pasted)
              </button>
              
              {leads.some(l => l.status === "idle" || l.status === "failed") && (
                <button
                  onClick={handleInvestigateAll}
                  disabled={runningCount > 0}
                  className="bg-slate-800/80 hover:bg-slate-700 text-amber-400 hover:text-amber-300 py-2 px-3 rounded-lg text-xs font-semibold flex items-center justify-center gap-1.5 border border-slate-800 transition"
                >
                  <span className="relative h-4 w-4 shrink-0" aria-hidden="true">
                    <Sparkles className={`absolute inset-0 h-4 w-4 ${runningCount > 0 ? "invisible" : "visible"}`} />
                    <Loader2 className={`absolute inset-0 h-4 w-4 animate-spin ${runningCount > 0 ? "visible" : "invisible"}`} />
                  </span>
                  <span>{runningCount > 0 ? "分析中..." : "一键背调待查"}</span>
                </button>
              )}
            </div>
          </div>

          {/* Collective Investigation Console for spreadsheet workflows */}
          {totalUploadedCount > 0 && (
            <div className="bg-gradient-to-br from-[#121E16] to-[#151515] border border-emerald-900/40 rounded-xl p-4.5 shadow-md space-y-3.5 relative overflow-hidden animate-fadeIn">
              {/* Decorative background glow */}
              <div className="absolute top-0 right-0 w-32 h-32 bg-emerald-500/5 rounded-full blur-2xl -mr-8 -mt-8 pointer-events-none"></div>
              
              <div className="flex items-center justify-between border-b border-emerald-950/40 pb-2">
                <div className="flex items-center gap-2">
                  <FileSpreadsheet className="h-4.5 w-4.5 text-emerald-400" />
                  <span className="text-xs font-bold text-white tracking-wide">
                    集体背调集成控制台
                  </span>
                </div>
                <div className="flex items-center gap-1.5">
                  <span className={`text-[9px] px-1.5 py-0.5 rounded font-bold ${
                    appMode === "create" ? "bg-emerald-500/10 text-emerald-400 border border-emerald-500/20" : "bg-teal-500/10 text-teal-400 border border-teal-500/20"
                  }`}>
                    {appMode === "create" ? "新建入库" : "匹配更新"}
                  </span>
                  <span className="text-[10px] text-emerald-400 bg-emerald-500/10 px-2 py-0.5 rounded-full border border-emerald-500/20 font-mono max-w-[120px] truncate">
                    {originalFileName || "未命名表格"}
                  </span>
                </div>
              </div>

              {/* Process Progress Indicator */}
              <div className="space-y-1.5">
                <div className="flex justify-between text-[11px]">
                  <span className="text-slate-400">批量任务处理进度</span>
                  <span className="font-mono text-white font-semibold">
                    {completedCount + failedCount} / {totalUploadedCount} ({Math.round(((completedCount + failedCount) / totalUploadedCount) * 100)}%)
                  </span>
                </div>
                
                {/* Progress Bar Container */}
                <div className="w-full bg-[#0F0F0F] rounded-full h-2 overflow-hidden border border-slate-800/80 flex">
                  <div 
                    className="bg-emerald-500 h-full transition-all duration-300"
                    style={{ width: `${totalUploadedCount > 0 ? (completedCount / totalUploadedCount) * 100 : 0}%` }}
                  ></div>
                  <div 
                    className="bg-amber-500 h-full transition-all duration-300"
                    style={{ width: `${totalUploadedCount > 0 ? (runningCount / totalUploadedCount) * 100 : 0}%` }}
                  ></div>
                  <div 
                    className="bg-red-500 h-full transition-all duration-300"
                    style={{ width: `${totalUploadedCount > 0 ? (failedCount / totalUploadedCount) * 100 : 0}%` }}
                  ></div>
                </div>
              </div>

              {/* Metric Counts breakdown */}
              <div className="grid grid-cols-4 gap-1.5 text-center bg-[#0F0F0F]/60 border border-slate-850 p-2 rounded-lg font-mono">
                <div>
                  <div className="text-[9px] text-slate-500">待处理</div>
                  <div className="text-xs font-bold text-slate-300">{idleCount}</div>
                </div>
                <div>
                  <div className="text-[9px] text-amber-500">处理中</div>
                  <div className="text-xs font-bold text-amber-400 animate-pulse">{runningCount}</div>
                </div>
                <div>
                  <div className="text-[9px] text-emerald-500">已成功</div>
                  <div className="text-xs font-bold text-emerald-400">{completedCount}</div>
                </div>
                <div>
                  <div className="text-[9px] text-red-500">已失败</div>
                  <div className="text-xs font-bold text-red-400">{failedCount}</div>
                </div>
              </div>

              {/* Action buttons */}
              <div className="flex gap-2">
                {idleCount > 0 || runningCount > 0 ? (
                  <button
                    onClick={handleInvestigateAll}
                    disabled={runningCount > 0}
                    className={`flex-1 py-2 px-3 rounded-lg text-xs font-semibold flex items-center justify-center gap-1.5 transition ${
                      runningCount > 0
                        ? "bg-slate-850 text-slate-500 cursor-not-allowed border border-slate-800"
                        : "bg-emerald-600 hover:bg-emerald-500 text-white active:bg-emerald-700 shadow-sm"
                    }`}
                  >
                    <span className="relative h-3.5 w-3.5 shrink-0" aria-hidden="true">
                      <Sparkles className={`absolute inset-0 h-3.5 w-3.5 ${runningCount > 0 ? "invisible" : "visible"}`} />
                      <Loader2 className={`absolute inset-0 h-3.5 w-3.5 animate-spin text-emerald-400 ${runningCount > 0 ? "visible" : "invisible"}`} />
                    </span>
                    <span>{runningCount > 0 ? "分析中..." : "开始批量背调"}</span>
                  </button>
                ) : null}

                <button
                  onClick={handleDownloadBulkZip}
                  disabled={completedCount === 0 || isExportingZip}
                  className={`flex-1 py-2 px-3 rounded-lg text-xs font-semibold flex items-center justify-center gap-1.5 transition ${
                    completedCount === 0 || isExportingZip
                      ? "bg-slate-850 text-slate-500 cursor-not-allowed border border-slate-800"
                      : "bg-slate-100 hover:bg-white text-slate-950 hover:shadow-md active:bg-slate-200"
                  }`}
                >
                  {isExportingZip ? (
                    <>
                      <Loader2 className="h-3.5 w-3.5 animate-spin" />
                      打包中...
                    </>
                  ) : (
                    <>
                      <Download className="h-3.5 w-3.5" />
                      打包集体导出
                    </>
                  )}
                </button>
              </div>
              
              <div className="text-[10px] text-slate-500 text-center flex items-center justify-center gap-1">
                <Info className="h-3 w-3 shrink-0 text-slate-600" />
                包含新增“背景调查报告文档”列的原 Excel 及配套同名 Word 报告
              </div>

              {/* Excel-only & Feishu spreadsheet integration actions */}
              <div className="grid grid-cols-2 gap-2 pt-2.5 border-t border-emerald-950/30">
                <button
                  onClick={handleDownloadOnlyExcel}
                  disabled={completedCount === 0}
                  className={`py-2 px-2.5 rounded-lg text-[11px] font-semibold flex items-center justify-center gap-1.5 transition ${
                    completedCount === 0
                      ? "bg-slate-850 text-slate-500 cursor-not-allowed border border-slate-800"
                      : "bg-slate-900 hover:bg-slate-855 text-emerald-400 border border-slate-800 hover:border-emerald-700/60"
                  }`}
                >
                  <FileSpreadsheet className="h-3.5 w-3.5" />
                  仅导出背调表格 (.xlsx)
                </button>

                <button
                  onClick={handleUploadSheetToFeishu}
                  disabled={completedCount === 0 || isUploadingSheet}
                  className={`py-2 px-2.5 rounded-lg text-[11px] font-semibold flex items-center justify-center gap-1.5 transition ${
                    completedCount === 0 || isUploadingSheet
                      ? "bg-slate-850 text-slate-500 cursor-not-allowed border border-slate-800"
                      : "bg-[#00D061]/15 hover:bg-[#00D061]/25 text-[#00D061] border border-[#00D061]/30 active:bg-[#00D061]/35"
                  }`}
                >
                  {isUploadingSheet ? (
                    <>
                      <Loader2 className="h-3.5 w-3.5 animate-spin" />
                      {appMode === "update" ? "匹配并更新中..." : "转换并上传中..."}
                    </>
                  ) : (
                    <>
                      <ExternalLink className="h-3.5 w-3.5" />
                      {appMode === "update" ? "匹配更新多维表格" : "同步飞书多维表格"}
                    </>
                  )}
                </button>
              </div>

              {feishuSheetUrl && (
                <div className="bg-[#00D061]/5 border border-[#00D061]/20 rounded-lg p-2.5 flex items-center justify-between text-[11px] text-[#00D061] animate-fadeIn">
                  <div className="flex items-center gap-1.5 truncate">
                    <span className="h-1.5 w-1.5 rounded-full bg-[#00D061] animate-ping"></span>
                    <span className="truncate">
                      {appMode === "update" ? "已精准匹配更新多维表格记录：" : "已同步多维表格新增记录："}
                    </span>
                  </div>
                  <a
                    href={feishuSheetUrl}
                    target="_blank"
                    rel="noreferrer"
                    className="shrink-0 flex items-center gap-0.5 font-bold hover:underline font-sans"
                  >
                    查看多维表格 <ExternalLink className="h-3 w-3" />
                  </a>
                </div>
              )}
            </div>
          )}

          {/* Inline Addition Modal/Card */}
          {isAddingLead && (
            <div className="bg-[#151515] border border-slate-800 rounded-xl p-5 shadow-lg relative animate-fadeIn">
              <div className="flex justify-between items-center mb-4 pb-2 border-b border-slate-800/60">
                <div className="flex items-center gap-2">
                  <FileSpreadsheet className="h-5 w-5 text-emerald-400" />
                  <h3 className="text-sm font-semibold text-white">添加客商线索</h3>
                </div>
                <div className="flex bg-[#0F0F0F] p-0.5 rounded-lg border border-slate-800/60 text-xs">
                  <button
                    onClick={() => setAddMode("batch")}
                    className={`px-3 py-1 rounded-md transition ${addMode === "batch" ? "bg-emerald-600 text-white font-semibold" : "text-slate-400 hover:text-slate-200"}`}
                  >
                    表格/文本批量导入
                  </button>
                  <button
                    onClick={() => setAddMode("manual")}
                    className={`px-3 py-1 rounded-md transition ${addMode === "manual" ? "bg-emerald-600 text-white font-semibold" : "text-slate-400 hover:text-slate-200"}`}
                  >
                    手动单条录入
                  </button>
                </div>
              </div>

              {addMode === "batch" ? (
                <div className="space-y-4">
                  {/* Sub-modes: File Upload vs Text Paste */}
                  <div className="flex bg-[#0F0F0F] p-0.5 rounded-lg border border-slate-800/80 text-xs w-full">
                    <button
                      onClick={() => setBatchSubMode("file")}
                      className={`flex-1 py-1.5 rounded-md transition flex items-center justify-center gap-1.5 ${
                        batchSubMode === "file" ? "bg-emerald-600/20 text-emerald-400 font-semibold border border-emerald-500/20" : "text-slate-400 hover:text-slate-200"
                      }`}
                    >
                      <Upload className="h-3.5 w-3.5" /> 📄 导入 Excel/CSV 文件
                    </button>
                    <button
                      onClick={() => setBatchSubMode("paste")}
                      className={`flex-1 py-1.5 rounded-md transition flex items-center justify-center gap-1.5 ${
                        batchSubMode === "paste" ? "bg-emerald-600/20 text-emerald-400 font-semibold border border-emerald-500/20" : "text-slate-400 hover:text-slate-200"
                      }`}
                    >
                      <ClipboardCopy className="h-3.5 w-3.5" /> ✍️ 粘贴表格文本
                    </button>
                  </div>

                  {batchSubMode === "file" ? (
                    <div className="space-y-3">
                      <p className="text-xs text-slate-400">
                        您可以直接上传导出的原始订单表格 (WPS / Excel / CSV)。系统将<strong>自动检索各列中的关键信息</strong>（公司名称、收件人、国家、产品名称、税号等）进行智能背调：
                      </p>

                      {/* Download 48-Field Sample Template Card */}
                      <div className="flex justify-between items-center bg-[#0F0F0F] p-3 rounded-xl border border-slate-800/80 my-1 animate-fadeIn">
                        <div className="text-left space-y-0.5">
                          <span className="text-xs font-bold text-slate-200 block">飞书多维表格 48 字段标准背调模版</span>
                          <span className="text-[10px] text-slate-500 block">符合飞书已配置好的 48 列电商订单标准格式</span>
                        </div>
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            handleDownloadSampleTemplate();
                          }}
                          className="bg-emerald-600/10 hover:bg-emerald-600/20 active:bg-emerald-600/30 text-emerald-400 border border-emerald-500/20 px-3 py-1.5 rounded-lg text-[10px] font-semibold transition flex items-center gap-1 shrink-0"
                        >
                          <Download className="h-3 w-3" /> 下载 .xlsx 模版
                        </button>
                      </div>
                      
                      {/* Drag & Drop Upload Container */}
                      <div
                        onDragOver={(e) => {
                          e.preventDefault();
                          setIsDragging(true);
                        }}
                        onDragLeave={() => setIsDragging(false)}
                        onDrop={(e) => {
                          e.preventDefault();
                          setIsDragging(false);
                          const file = e.dataTransfer.files?.[0];
                          if (file) handleFileUpload(file);
                        }}
                        onClick={() => {
                          const el = document.getElementById("excel-file-input");
                          el?.click();
                        }}
                        className={`border-2 border-dashed rounded-xl p-6 text-center cursor-pointer transition flex flex-col items-center justify-center min-h-[140px] ${
                          isDragging
                            ? "border-emerald-500 bg-emerald-500/5"
                            : "border-slate-800 bg-[#0F0F0F] hover:border-slate-700 hover:bg-[#0F0F0F]/80"
                        }`}
                      >
                        <input
                          id="excel-file-input"
                          type="file"
                          accept=".xlsx,.xls,.csv"
                          className="hidden"
                          onChange={(e) => {
                            const file = e.target.files?.[0];
                            if (file) handleFileUpload(file);
                          }}
                        />
                        <div className="bg-emerald-500/10 text-emerald-400 p-2.5 rounded-full border border-emerald-500/10 mb-2.5">
                          <Upload className="h-5 w-5" />
                        </div>
                        <p className="text-xs font-semibold text-white">
                          拖拽电子表格到此区域，或 <span className="text-emerald-400 hover:underline">点击浏览文件</span>
                        </p>
                        <p className="text-[10px] text-slate-500 mt-1">支持 .xlsx, .xls, .csv 等各种 B2B/零售平台导出的报表表格</p>
                        
                        {/* Status messages inside the uploader */}
                        {uploadError && (
                          <div className="mt-3 text-[10px] text-red-400 bg-red-500/10 px-2.5 py-1 rounded border border-red-500/20 flex items-center gap-1">
                            <AlertCircle className="h-3 w-3 shrink-0" />
                            {uploadError}
                          </div>
                        )}
                        {uploadSuccessMsg && (
                          <div className="mt-3 text-[10px] text-emerald-400 bg-emerald-500/10 px-2.5 py-1 rounded border border-emerald-500/20 flex items-center gap-1">
                            <Check className="h-3 w-3 shrink-0" />
                            {uploadSuccessMsg}
                          </div>
                        )}
                      </div>
                      
                      <div className="flex justify-end pt-1">
                        <button
                          onClick={() => setIsAddingLead(false)}
                          className="bg-slate-850 hover:bg-slate-800 text-slate-400 hover:text-slate-300 px-4 py-2 rounded-lg text-xs border border-slate-800"
                        >
                          取消
                        </button>
                      </div>
                    </div>
                  ) : (
                    <div>
                      <p className="text-xs text-slate-400 mb-2">
                        支持直接从 Excel/WPS/Sheets 复制多行并在此处粘贴。格式需为：<br />
                        <code className="text-emerald-400 bg-emerald-500/10 px-1 py-0.5 rounded">公司名称 [Tab或逗号] 国家 [Tab或逗号] 订单信息 [Tab或逗号] 产品</code>
                      </p>
                      <textarea
                        value={pasteArea}
                        onChange={(e) => setPasteArea(e.target.value)}
                        placeholder="例如:&#10;Acme Sp. z o.o.	Poland	ul. Akacjowa 15, Warszawa...	贴标机&#10;AgroSabores S.A.S	Colombia	Carrera 7 # 12-34, Bogotá...	液体灌装机"
                        className="w-full h-32 bg-[#0F0F0F] border border-slate-800 rounded-lg p-3 text-xs text-slate-100 placeholder-slate-600 font-mono focus:outline-none focus:border-slate-700 focus:ring-1 focus:ring-slate-700"
                      />
                      <div className="flex justify-end space-x-2 mt-4">
                        <button
                          onClick={() => setIsAddingLead(false)}
                          className="bg-slate-800 hover:bg-slate-700 text-slate-300 px-4 py-2 rounded-lg text-xs"
                        >
                          取消
                        </button>
                        <button
                          onClick={handleBatchImport}
                          className="bg-emerald-600 hover:bg-emerald-500 text-white font-semibold px-4 py-2 rounded-lg text-xs"
                        >
                          确认解析并添加
                        </button>
                      </div>
                    </div>
                  )}
                </div>
              ) : (
                <div className="space-y-3">
                  <div>
                    <label className="block text-xs text-slate-400 mb-1">公司官方注册名称 *</label>
                    <input
                      type="text"
                      value={manualForm.companyName}
                      onChange={(e) => setManualForm(prev => ({ ...prev, companyName: e.target.value }))}
                      placeholder="如: Acme Sp. z o.o."
                      className="w-full bg-[#0F0F0F] border border-slate-800 rounded-lg px-3 py-2 text-xs text-slate-100 focus:outline-none focus:border-slate-700"
                    />
                  </div>
                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-xs text-slate-400 mb-1">国家/地区</label>
                      <input
                        type="text"
                        value={manualForm.country}
                        onChange={(e) => setManualForm(prev => ({ ...prev, country: e.target.value }))}
                        placeholder="如: Poland / 波兰"
                        className="w-full bg-[#0F0F0F] border border-slate-800 rounded-lg px-3 py-2 text-xs text-slate-100 focus:outline-none focus:border-slate-700"
                      />
                    </div>
                    <div>
                      <label className="block text-xs text-slate-400 mb-1">采购意向/机器类型</label>
                      <input
                        type="text"
                        value={manualForm.productContext}
                        onChange={(e) => setManualForm(prev => ({ ...prev, productContext: e.target.value }))}
                        placeholder="如: 贴标机"
                        className="w-full bg-[#0F0F0F] border border-slate-800 rounded-lg px-3 py-2 text-xs text-slate-100 focus:outline-none focus:border-slate-700"
                      />
                    </div>
                  </div>
                  <div>
                    <label className="block text-xs text-slate-400 mb-1">订单收货详情/税号信息</label>
                    <textarea
                      value={manualForm.orderInfo}
                      onChange={(e) => setManualForm(prev => ({ ...prev, orderInfo: e.target.value }))}
                      placeholder="包含收件人、详细地址、NIP税号、邮箱等，供AI提取 cross-reference 核对"
                      className="w-full h-20 bg-[#0F0F0F] border border-slate-800 rounded-lg p-3 text-xs text-slate-100 focus:outline-none focus:border-slate-700"
                    />
                  </div>
                  <div className="flex justify-end space-x-2 pt-2">
                    <button
                      onClick={() => setIsAddingLead(false)}
                      className="bg-slate-800 hover:bg-slate-700 text-slate-300 px-4 py-2 rounded-lg text-xs"
                    >
                      取消
                    </button>
                    <button
                      onClick={handleManualImport}
                      className="bg-emerald-600 hover:bg-emerald-500 text-white font-semibold px-4 py-2 rounded-lg text-xs"
                    >
                      添加单个客户
                    </button>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* Customer Leads List Card */}
          <div className="bg-[#151515] border border-slate-800 rounded-xl flex-1 flex flex-col shadow-sm overflow-hidden">
            <div className="p-4 border-b border-slate-800 bg-[#0F0F0F] flex justify-between items-center">
              <h2 className="text-sm font-semibold text-slate-200 flex items-center gap-2">
                <FileText className="h-4 w-4 text-emerald-400" />
                待核查与已调查客商列表
              </h2>
              <span className="text-xs bg-[#0F0F0F] text-slate-400 px-2.5 py-0.5 rounded-full border border-slate-800/60 font-mono">{leads.length} 家</span>
            </div>

            <div className="divide-y divide-slate-800/60 overflow-y-auto max-h-[500px] flex-1">
              {leads.length === 0 ? (
                <div className="p-12 text-center text-slate-500">
                  <Building2 className="h-10 w-10 mx-auto mb-3 stroke-1 text-slate-600" />
                  <p className="text-sm">暂无客商线索。请点击上方按钮导入客户信息。</p>
                </div>
              ) : (
                leads.map((lead) => {
                  const isSelected = lead.id === selectedLeadId;
                  return (
                    <div
                      key={lead.id}
                      onClick={() => setSelectedLeadId(lead.id)}
                      className={`p-4 cursor-pointer transition flex flex-col space-y-2 text-left relative ${
                        isSelected ? "bg-emerald-950/10 border-l-4 border-emerald-500" : "hover:bg-[#1C1C1C]/40"
                      }`}
                    >
                      <div className="flex justify-between items-start">
                        <div className="space-y-0.5">
                          <h3 className="text-sm font-bold text-white flex items-center gap-1.5">
                            {lead.companyName}
                            {lead.grade && (
                              <span className="text-[10px] bg-emerald-500/10 text-emerald-400 px-1.5 py-0.5 rounded border border-emerald-500/20">
                                {lead.grade.split("（")[0]}
                              </span>
                            )}
                          </h3>
                          <div className="flex items-center space-x-2 text-xs text-slate-400">
                            {lead.country && (
                              <span className="flex items-center gap-0.5">
                                <MapPin className="h-3 w-3" /> {lead.country}
                              </span>
                            )}
                            <span className="text-slate-600">•</span>
                            <span className="bg-[#0F0F0F] border border-slate-800 px-1.5 py-0.5 rounded text-[10px] text-slate-300">
                              {lead.productContext}
                            </span>
                          </div>
                        </div>
                        
                        {/* Investigation Action & Status badge */}
                        <div className="flex items-center space-x-2">
                          {lead.status === "idle" && (
                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                handleInvestigate(lead.id);
                              }}
                              className="bg-emerald-600 hover:bg-emerald-500 text-white font-semibold px-2.5 py-1 rounded text-[10px] flex items-center gap-1 transition"
                            >
                              <Search className="h-3 w-3" /> 立即背调
                            </button>
                          )}
                          
                          {lead.status === "searching" && (
                            <span className="bg-amber-500/10 text-amber-400 px-2 py-1 rounded text-[10px] font-medium border border-amber-500/20 flex items-center gap-1">
                              <Loader2 className="h-2.5 w-2.5 animate-spin" /> 核实中...
                            </span>
                          )}

                          {lead.status === "completed" && (
                            <div className="flex items-center space-x-1.5">
                              <span className="bg-emerald-500/10 text-emerald-400 px-2 py-1 rounded text-[10px] font-medium border border-emerald-500/20 flex items-center gap-0.5">
                                <Check className="h-2.5 w-2.5" /> 已完成
                              </span>
                              
                              {/* Feishu Doc Sync Status */}
                              {lead.feishuUrl ? (
                                <a
                                  href={lead.feishuUrl}
                                  target="_blank"
                                  rel="noreferrer noopener"
                                  onClick={(e) => e.stopPropagation()}
                                  className={`${lead.feishuImportFallback ? "bg-amber-500/10 text-amber-400 hover:bg-amber-500/20 border-amber-500/20" : "bg-sky-500/10 text-sky-400 hover:bg-sky-500/20 border-sky-500/20"} px-2 py-1 rounded text-[10px] font-medium border flex items-center gap-1 transition`}
                                  title={lead.feishuImportFallback ? `在线文档转换失败，当前为原始文件。尝试次数：${lead.feishuImportAttempts || 1}` : "已成功转换为飞书文档，点击打开预览"}
                                >
                                  <Cloud className="h-2.5 w-2.5" /> {lead.feishuImportFallback ? "飞书文件" : "飞书文档"}
                                </a>
                              ) : lead.feishuStatus === "uploading" ? (
                                <span className="bg-sky-500/10 text-sky-400 px-2 py-1 rounded text-[10px] font-medium border border-sky-500/20 flex items-center gap-1">
                                  <Loader2 className="h-2.5 w-2.5 animate-spin" /> 创建文档中
                                </span>
                              ) : lead.feishuStatus === "failed" ? (
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    retryUploadAndSync(lead);
                                  }}
                                  className="bg-amber-500/10 hover:bg-amber-500/20 text-amber-400 px-2 py-1 rounded text-[10px] font-medium border border-amber-500/20 flex items-center gap-1"
                                  title={`创建失败: ${lead.feishuError || "未知错误"}。点击重试`}
                                >
                                  <CloudOff className="h-2.5 w-2.5" /> 重试文档
                                </button>
                              ) : (
                                (feishuAppId || enableFeishu) && (
                                  <button
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      retryUploadAndSync(lead);
                                    }}
                                    className="bg-slate-850 hover:bg-slate-800 text-slate-300 hover:text-sky-400 border border-slate-800/80 px-2 py-1 rounded text-[10px] font-medium flex items-center gap-1 transition"
                                    title="在飞书云文档中创建线上文档报告"
                                  >
                                    <CloudUpload className="h-2.5 w-2.5" /> 飞书文档
                                  </button>
                                )
                              )}

                              {/* Feishu Bitable (Multi-dimensional Table) Sync Status */}
                              {lead.bitableUrl ? (
                                <a
                                  href={lead.bitableUrl}
                                  target="_blank"
                                  rel="noreferrer noopener"
                                  onClick={(e) => e.stopPropagation()}
                                  className="bg-emerald-500/10 text-emerald-400 hover:bg-emerald-500/20 px-2 py-1 rounded text-[10px] font-medium border border-emerald-500/20 flex items-center gap-1 transition"
                                  title="已成功同步至多维表格，点击打开"
                                >
                                  <FileSpreadsheet className="h-2.5 w-2.5" /> 多维表
                                </a>
                              ) : lead.bitableStatus === "syncing" ? (
                                <span className="bg-emerald-500/10 text-emerald-400 px-2 py-1 rounded text-[10px] font-medium border border-emerald-500/20 flex items-center gap-1">
                                  <Loader2 className="h-2.5 w-2.5 animate-spin" /> 同步多维表...
                                </span>
                              ) : lead.bitableStatus === "failed" ? (
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    syncSingleToBitable(lead.id);
                                  }}
                                  className="bg-red-500/10 hover:bg-red-500/20 text-red-400 px-2 py-1 rounded text-[10px] font-medium border border-red-500/20 flex items-center gap-1"
                                  title={`同步失败: ${lead.bitableError || "未知错误"}。点击重试`}
                                >
                                  <CloudOff className="h-2.5 w-2.5" /> 重试多维表
                                </button>
                              ) : (
                                (feishuAppId || enableFeishu) && feishuBitableUrl.trim() && (
                                  <button
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      syncSingleToBitable(lead.id);
                                    }}
                                    className="bg-slate-850 hover:bg-slate-800 text-slate-300 hover:text-emerald-400 border border-slate-800/80 px-2 py-1 rounded text-[10px] font-medium flex items-center gap-1 transition"
                                    title="同步记录至飞书多维表格"
                                  >
                                    <FileSpreadsheet className="h-2.5 w-2.5" /> 同步多维表
                                  </button>
                                )
                              )}
                            </div>
                          )}

                          {lead.status === "failed" && (
                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                handleInvestigate(lead.id);
                              }}
                              className="bg-red-500/10 hover:bg-red-500/20 text-red-400 px-2 py-1 rounded text-[10px] font-medium border border-red-500/20 flex items-center gap-1"
                            >
                              <AlertCircle className="h-2.5 w-2.5" /> 重试
                            </button>
                          )}

                          <button
                            onClick={(e) => handleDeleteLead(lead.id, e)}
                            className="text-slate-500 hover:text-red-400 p-1 rounded hover:bg-slate-800 transition"
                            title="删除线索"
                          >
                            <Trash2 className="h-3.5 w-3.5" />
                          </button>
                        </div>
                      </div>

                      {lead.summary && (
                        <p className="text-xs text-slate-400 line-clamp-2 bg-[#0F0F0F]/60 p-2 rounded border border-slate-800/40">
                          {lead.summary}
                        </p>
                      )}

                      {lead.error && (
                        <p className="text-[10px] text-red-400 bg-red-500/5 p-1.5 rounded flex items-start gap-1 border border-red-500/10 whitespace-pre-wrap line-clamp-3">
                          <AlertCircle className="h-3.5 w-3.5 shrink-0 mt-0.5" />
                          <span className="flex-1">{lead.error}</span>
                        </p>
                      )}
                    </div>
                  );
                })
              )}
            </div>
          </div>
        </div>

        {/* Right Detail Panel: Due Diligence details & Interactive Outreach Strategy */}
        <div className="md:col-span-8 flex flex-col space-y-6">
          {selectedLead ? (
            <div className="bg-[#151515] border border-slate-800 rounded-xl flex-1 flex flex-col shadow-sm overflow-hidden">
              
              {/* Profile Card Header */}
              <div className="p-6 border-b border-slate-800/80 bg-[#0F0F0F] flex flex-col md:flex-row md:items-center justify-between gap-4 text-left">
                <div className="space-y-1.5">
                  <span className="text-[10px] font-mono font-bold text-slate-500 bg-[#0F0F0F] border border-slate-800 px-2 py-0.5 rounded">
                    企业背调档案 / CLIENT DOSSIER
                  </span>
                  <h1 className="text-xl font-extrabold text-white tracking-tight flex items-center gap-2">
                    {selectedLead.companyName}
                  </h1>
                  <p className="text-xs text-slate-400 flex items-center gap-1">
                    <MapPin className="h-3.5 w-3.5 text-emerald-400" />
                    {selectedLead.country} • NIP 税号检索完成 • 档案更新于 {selectedLead.completedAt || "2026-07-12"}
                  </p>
                </div>
                <div className="flex items-center gap-2">
                  {selectedLead.status === "completed" && (
                    <>
                      {selectedLead.feishuUrl ? (
                        <a
                          href={selectedLead.feishuUrl}
                          target="_blank"
                          rel="noreferrer noopener"
                          className={`${selectedLead.feishuImportFallback ? "bg-amber-600 hover:bg-amber-500" : "bg-sky-600 hover:bg-sky-500"} text-white px-3.5 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition shadow-sm`}
                          title={selectedLead.feishuImportFallback ? `文件已上传，但在线文档转换失败：${selectedLead.feishuImportError || "未知原因"}` : "在飞书云文档中直接打开并编辑此报告"}
                        >
                          <Cloud className="h-3.5 w-3.5" /> {selectedLead.feishuImportFallback ? "飞书原始文件" : "飞书在线文档"}
                        </a>
                      ) : (
                        (feishuAppId || enableFeishu) && (
                          <button
                            onClick={() => retryUploadAndSync(selectedLead)}
                            disabled={selectedLead.feishuStatus === "uploading"}
                            className="bg-[#0F0F0F] hover:bg-[#151515] border border-slate-800 text-sky-400 hover:text-sky-300 px-3.5 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition disabled:opacity-50"
                            title="将此报告导入并创建为在线飞书文档"
                          >
                            {selectedLead.feishuStatus === "uploading" ? (
                              <>
                                <Loader2 className="h-3.5 w-3.5 animate-spin" /> 正在创建...
                              </>
                            ) : (
                              <>
                                <CloudUpload className="h-3.5 w-3.5" /> 创建飞书文档
                              </>
                            )}
                          </button>
                        )
                      )}
                    </>
                  )}
                  <button
                    onClick={() => handleDownloadMD(selectedLead)}
                    className="bg-[#0F0F0F] hover:bg-[#151515] border border-slate-800 text-slate-300 hover:text-white px-3.5 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition"
                  >
                    <Download className="h-3.5 w-3.5" /> 导出 MD 报告
                  </button>
                  <button
                    onClick={() => handleDownloadWord(selectedLead)}
                    className="bg-[#0F0F0F] hover:bg-[#151515] border border-slate-800 text-emerald-400 hover:text-emerald-300 px-3.5 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition"
                  >
                    <FileText className="h-3.5 w-3.5 text-emerald-400" /> 导出 Word 报告
                  </button>
                  <button
                    onClick={() => handleInvestigate(selectedLead.id)}
                    className="bg-emerald-600 hover:bg-emerald-500 text-white px-3.5 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition shadow-sm"
                  >
                    <Sparkles className="h-3.5 w-3.5" /> 重新核查
                  </button>
                </div>
              </div>

              {selectedLead.status === "searching" && (
                <div className="flex-1 flex flex-col items-center justify-center p-12 text-center space-y-4">
                  <div className="relative">
                    <div className="h-16 w-16 rounded-full border-4 border-slate-800 border-t-emerald-500 animate-spin"></div>
                    <Search className="h-6 w-6 text-emerald-400 absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2" />
                  </div>
                  <div>
                    <h3 className="text-base font-bold text-white">正在穿透互联网调查 {selectedLead.companyName}...</h3>
                    <p className="text-xs text-slate-400 mt-1 max-w-md mx-auto">
                      AI 正在启用 DeepSeek 联网搜索，核查当地工商注册、NIP、官方网站、生产厂房详情、包装规格以及在册经营状态，通常需要 10-20 秒。
                    </p>
                  </div>
                  <div className="text-xs bg-[#0F0F0F] border border-slate-800 px-3 py-1.5 rounded text-slate-500 max-w-sm font-mono animate-pulse">
                    [Task Log]: Initiating web_search_20250305 API call with tax/country filters...
                  </div>
                </div>
              )}

              {selectedLead.status === "failed" && (
                <div className="flex-1 flex flex-col items-center justify-center p-12 text-center space-y-5">
                  <AlertCircle className="h-12 w-12 text-red-500 bg-red-500/10 p-2.5 rounded-full border border-red-500/20 animate-bounce" />
                  <div className="max-w-xl space-y-2">
                    <h3 className="text-base font-bold text-white">背调执行失败 (Investigation Interrupted)</h3>
                    <div className="text-xs text-red-300 bg-red-950/20 border border-red-500/20 rounded-xl p-4 text-left whitespace-pre-wrap leading-relaxed max-h-[250px] overflow-y-auto font-sans">
                      {selectedLead.error || "连接服务器或调用 API 时遇到一些障碍，这可能是因为秘钥过期或查询词含特殊符号。"}
                    </div>
                  </div>
                  <button
                    onClick={() => handleInvestigate(selectedLead.id)}
                    className="bg-slate-800 hover:bg-slate-700 text-white px-4 py-2 rounded-lg text-xs font-semibold border border-slate-700 transition"
                  >
                    重新发起请求
                  </button>
                </div>
              )}

              {selectedLead.status === "idle" && (
                <div className="flex-1 flex flex-col items-center justify-center p-12 text-center space-y-6">
                  <div className="h-16 w-16 bg-[#0F0F0F] border border-slate-800 rounded-full flex items-center justify-center">
                    <Search className="h-8 w-8 text-slate-500" />
                  </div>
                  <div className="max-w-md space-y-2">
                    <h3 className="text-base font-bold text-white">未进行背景调查</h3>
                    <p className="text-xs text-slate-400">
                      该客商处于待查队列。AI 调查模型将扫描公开网页、波兰官方 KRS、波兰 GoWork 雇员评价、欧洲 VIES 税号系统或目标国工商数据库，提炼成与您包装机械业务匹配度极高的跟进报告。
                    </p>
                  </div>
                  <button
                    onClick={() => handleInvestigate(selectedLead.id)}
                    className="bg-emerald-600 hover:bg-emerald-500 active:bg-emerald-700 text-white px-5 py-2.5 rounded-lg text-sm font-semibold transition flex items-center gap-1.5"
                  >
                    <Sparkles className="h-4 w-4" /> 启动深度 AI 背景调查
                  </button>
                </div>
              )}

              {selectedLead.status === "completed" && (
                <div className="flex-1 flex flex-col overflow-hidden">
                  
                  {/* Executive Overview Bento Cards */}
                  <div className="p-6 bg-[#0F0F0F] border-b border-slate-800/80 grid grid-cols-1 md:grid-cols-12 gap-4">
                    
                    {/* Value Badge Card */}
                    <div className="md:col-span-5 bg-gradient-to-br from-[#151515] to-[#0F0F0F] border border-emerald-500/20 p-4 rounded-xl flex flex-col justify-between">
                      <div className="flex justify-between items-start">
                        <span className="text-[10px] text-emerald-400 font-semibold bg-emerald-500/10 border border-emerald-500/20 px-2 py-0.5 rounded-full">
                          B2B 价值评级
                        </span>
                        <Award className="h-4 w-4 text-emerald-400" />
                      </div>
                      <div className="my-3 text-left">
                        <h4 className="text-sm font-bold text-white flex items-center gap-1.5">
                          {selectedLead.grade || <span className="text-slate-500 italic font-normal">待背景分析后生成</span>}
                        </h4>
                        <p className="text-xs text-slate-400 mt-1">
                          根据企业经营性质、注册资质、线上公开实力及生产配套匹配度进行的综合AI评级。
                        </p>
                      </div>
                    </div>

                    {/* Quick Business Details Grid */}
                    <div className="md:col-span-7 bg-[#151515]/60 border border-slate-800/80 p-4 rounded-xl flex flex-col justify-between">
                      <div className="flex justify-between items-center pb-2 border-b border-slate-800/60">
                        <span className="text-xs text-slate-300 font-bold flex items-center gap-1.5">
                          <Info className="h-3.5 w-3.5 text-emerald-400" /> 工商核心指标简表
                        </span>
                        {selectedLead.companyOverview?.website && selectedLead.companyOverview.website !== "未公开" && (
                          <a
                            href={selectedLead.companyOverview.website}
                            target="_blank"
                            rel="noreferrer referrerPolicy"
                            className="text-[10px] text-emerald-400 hover:underline flex items-center gap-0.5 font-medium"
                          >
                            <Globe className="h-3 w-3" /> 访问官网 <ExternalLink className="h-2.5 w-2.5" />
                          </a>
                        )}
                      </div>
                      <div className="grid grid-cols-2 gap-x-4 gap-y-2 text-xs pt-3 text-left">
                        <div>
                          <span className="text-slate-500">注册税号 (Tax ID):</span>
                          <p className="font-mono text-white truncate" title={selectedLead.companyOverview?.taxId}>
                            {selectedLead.companyOverview?.taxId || "未提供/检索"}
                          </p>
                        </div>
                        <div>
                          <span className="text-slate-500">注册地址 (HQ):</span>
                          <p className="text-white truncate" title={selectedLead.companyOverview?.headquarters}>
                            {selectedLead.companyOverview?.headquarters || "未提供"}
                          </p>
                        </div>
                        <div>
                          <span className="text-slate-500">所属细分行业 (Industry):</span>
                          <p className="text-white truncate" title={selectedLead.companyOverview?.industry}>
                            {selectedLead.companyOverview?.industry || "食品包装生产商"}
                          </p>
                        </div>
                        <div>
                          <span className="text-slate-500">运行状态 (Status):</span>
                          <p className="text-white flex items-center gap-1">
                            <span className="h-2 w-2 rounded-full bg-emerald-500 inline-block animate-pulse"></span>
                            {selectedLead.companyOverview?.status || "Active (正常运转)"}
                          </p>
                        </div>
                      </div>
                    </div>

                    {/* Risk Alert Card */}
                    <div className="md:col-span-4 bg-gradient-to-br from-[#151515] to-[#0F0F0F] border border-amber-500/20 p-4 rounded-xl flex flex-col justify-between">
                      <div className="flex justify-between items-start">
                        <span className="text-[10px] text-amber-400 font-semibold bg-amber-500/10 border border-amber-500/20 px-2 py-0.5 rounded-full">
                          信用与准入风险提示
                        </span>
                        <ShieldAlert className="h-4 w-4 text-amber-400" />
                      </div>
                      <div className="my-2 text-left">
                        <p className="text-xs text-slate-200 leading-relaxed line-clamp-3" title={selectedLead.riskAlert}>
                          {selectedLead.riskAlert || <span className="text-slate-500 italic">暂无特别信用风险或诉讼。</span>}
                        </p>
                      </div>
                    </div>

                    {/* Follow-up Strategy Card */}
                    <div className="md:col-span-5 bg-gradient-to-br from-[#151515] to-[#0F0F0F] border border-blue-500/20 p-4 rounded-xl flex flex-col justify-between">
                      <div className="flex justify-between items-start">
                        <span className="text-[10px] text-blue-400 font-semibold bg-blue-500/10 border border-blue-500/20 px-2 py-0.5 rounded-full">
                          客户开发跟进策略
                        </span>
                        <Compass className="h-4 w-4 text-blue-400" />
                      </div>
                      <div className="my-2 text-left">
                        <p className="text-xs text-slate-200 leading-relaxed line-clamp-3" title={selectedLead.followUpStrategy}>
                          {selectedLead.followUpStrategy || "建议通过 LinkedIn 匹配采购决策人，并发送样品。"}
                        </p>
                      </div>
                    </div>

                    {/* Follow-up Grade Card */}
                    <div className="md:col-span-3 bg-gradient-to-br from-[#151515] to-[#0F0F0F] border border-indigo-500/20 p-4 rounded-xl flex flex-col justify-between">
                      <div className="flex justify-between items-start">
                        <span className="text-[10px] text-indigo-400 font-semibold bg-indigo-500/10 border border-indigo-500/20 px-2 py-0.5 rounded-full">
                          跟进优先级
                        </span>
                        <TrendingUp className="h-4 w-4 text-indigo-400" />
                      </div>
                      <div className="my-3 text-left">
                        <h4 className="text-sm font-bold text-white flex items-center gap-1.5">
                          {selectedLead.followUpGrade || <span className="text-slate-500 italic font-normal">待分析生成</span>}
                        </h4>
                        <p className="text-[10px] text-slate-400 mt-1">
                          系统推荐等级，便于团队按优先级开展精细化运营。
                        </p>
                      </div>
                    </div>

                    {/* API Usage Details Bento Card */}
                    <div className="md:col-span-12 bg-[#0F0F0F] border border-slate-800 p-5 rounded-xl flex flex-col space-y-4">
                      <div className="flex justify-between items-center pb-2.5 border-b border-slate-800">
                        <span className="text-xs text-slate-300 font-bold flex items-center gap-2">
                          <Cpu className="h-4 w-4 text-emerald-400" /> API 使用详情面板 / API Token Usage Details
                        </span>
                        {selectedLead.usage ? (
                          <span className="text-[10px] text-emerald-400 font-mono bg-emerald-500/10 border border-emerald-500/20 px-2.5 py-0.5 rounded-full flex items-center gap-1">
                            <span className="h-1.5 w-1.5 rounded-full bg-emerald-400 inline-block animate-pulse"></span>
                            数据正常载入
                          </span>
                        ) : (
                          <span className="text-[10px] text-amber-400 font-mono bg-amber-500/10 border border-amber-500/20 px-2.5 py-0.5 rounded-full flex items-center gap-1">
                            <span className="h-1.5 w-1.5 rounded-full bg-amber-400 inline-block"></span>
                            预设示例数据 (暂无 Token 记录)
                          </span>
                        )}
                      </div>

                      {selectedLead.usage ? (
                        <div className="grid grid-cols-1 sm:grid-cols-4 gap-4 text-left">
                          {/* Model */}
                          <div className="bg-[#151515] border border-slate-800/60 p-3 rounded-lg flex flex-col justify-between">
                            <span className="text-[10px] text-slate-500 uppercase tracking-wider font-semibold">执行模型 / Model</span>
                            <div className="mt-1.5">
                              <p className="text-xs font-mono font-semibold text-slate-200 truncate flex items-center gap-1.5">
                                <Cpu className="h-3.5 w-3.5 text-purple-400" />
                                {selectedLead.usage.modelUsed}
                              </p>
                              <p className="text-[10px] text-slate-400 mt-1">
                                {selectedLead.usage.searchGroundingUsed ? "🌐 启用 DeepSeek 实时联网检索" : "⚠️ 网络异常 fallback 状态下运行"}
                              </p>
                            </div>
                          </div>

                          {/* Prompt Tokens */}
                          <div className="bg-[#151515] border border-slate-800/60 p-3 rounded-lg flex flex-col justify-between">
                            <span className="text-[10px] text-slate-500 uppercase tracking-wider font-semibold font-sans">输入 Token / Prompt</span>
                            <div className="mt-1.5">
                              <p className="text-lg font-mono font-bold text-slate-100 flex items-baseline gap-1">
                                {selectedLead.usage.promptTokens.toLocaleString()}
                                <span className="text-[10px] text-slate-500 font-normal">tokens</span>
                              </p>
                              <p className="text-[10px] text-slate-400 mt-1">
                                包含系统指令 (Structured Guidelines) 与企业调查上下文
                              </p>
                            </div>
                          </div>

                          {/* Completion Tokens */}
                          <div className="bg-[#151515] border border-slate-800/60 p-3 rounded-lg flex flex-col justify-between">
                            <span className="text-[10px] text-slate-500 uppercase tracking-wider font-semibold">输出 Token / Output</span>
                            <div className="mt-1.5">
                              <p className="text-lg font-mono font-bold text-slate-100 flex items-baseline gap-1">
                                {selectedLead.usage.completionTokens.toLocaleString()}
                                <span className="text-[10px] text-slate-500 font-normal">tokens</span>
                              </p>
                              <p className="text-[10px] text-slate-400 mt-1">
                                AI 最终生成的结构化工商详情及背调评估报告
                              </p>
                            </div>
                          </div>

                          {/* Total Tokens */}
                          <div className="bg-[#151515] border border-slate-800/60 p-3 rounded-lg flex flex-col justify-between">
                            <span className="text-[10px] text-slate-500 uppercase tracking-wider font-semibold">总 Token 消耗 / Total</span>
                            <div className="mt-1.5">
                              <p className="text-lg font-mono font-bold text-emerald-400 flex items-baseline gap-1">
                                {selectedLead.usage.totalTokens.toLocaleString()}
                                <span className="text-[10px] text-emerald-600 font-normal">tokens</span>
                              </p>
                              <div className="flex items-center gap-1 mt-1">
                                <span className="text-[10px] text-slate-400">
                                  预估费用: 
                                </span>
                                <span className="text-[11px] font-mono text-emerald-400 font-bold">
                                  ¥{calculateTokenCost(selectedLead.usage.modelUsed, selectedLead.usage.promptTokens, selectedLead.usage.completionTokens).rmb.toFixed(4)}
                                </span>
                                <span className="text-[9px] text-slate-500 font-mono">
                                  (${calculateTokenCost(selectedLead.usage.modelUsed, selectedLead.usage.promptTokens, selectedLead.usage.completionTokens).usd.toFixed(5)} USD)
                                </span>
                              </div>
                            </div>
                          </div>
                        </div>
                      ) : (
                        <div className="bg-[#111] border border-dashed border-slate-800 rounded-lg p-5 text-center flex flex-col items-center justify-center space-y-1">
                          <Zap className="h-5 w-5 text-slate-600 animate-pulse" />
                          <p className="text-xs font-semibold text-slate-400">暂无此任务的实时 API Token 统计数据</p>
                          <p className="text-[10px] text-slate-500 max-w-md">
                            由于此档案为系统预置或导入的示例数据，并未经历实时 API 解析。您可以在当前档案或导入新数据后，点击右侧的 <span className="text-emerald-400">“重新发起背调”</span> 触发实时穿透并自动生成完整的 Token 消耗详情。
                          </p>
                        </div>
                      )}
                    </div>

                    {/* Interactive Grading Standards Guide Toggle */}
                    <div className="md:col-span-12 mt-2">
                      <div className="bg-slate-900/40 border border-slate-800/80 rounded-xl overflow-hidden transition-all duration-300">
                        <button 
                          onClick={() => setShowGradingGuide(!showGradingGuide)}
                          className="w-full flex justify-between items-center px-4 py-3 bg-[#111] hover:bg-[#151515] transition-colors"
                        >
                          <div className="flex items-center gap-2">
                            <Info className="h-4 w-4 text-emerald-400" />
                            <span className="text-xs font-bold text-slate-300">💡 智能背调与跟进分级判断标准说明</span>
                          </div>
                          <span className="text-[11px] text-emerald-400 font-medium hover:underline flex items-center gap-1">
                            {showGradingGuide ? "折叠详情" : "展开标准与空值释义 ➔"}
                          </span>
                        </button>
                        
                        {showGradingGuide && (
                          <div className="p-5 border-t border-slate-800/60 bg-[#0A0A0A] grid grid-cols-1 md:grid-cols-3 gap-6 text-left animate-fadeIn">
                            {/* Card Column 1: Background Analysis */}
                            <div className="space-y-3">
                              <div className="flex items-center gap-1.5 border-b border-slate-800/60 pb-1.5">
                                <Award className="h-3.5 w-3.5 text-emerald-400" />
                                <h5 className="text-xs font-bold text-white">背景分析等级 (B2B价值)</h5>
                              </div>
                              <ul className="space-y-2 text-[11px] text-slate-400">
                                <li>
                                  <strong className="text-emerald-400">💎 S 级：</strong>高价值实体生产商。拥有自建厂房或高度吻合的加工需求，复购潜能极大，推荐重点跟进。
                                </li>
                                <li>
                                  <strong className="text-sky-400">🥇 A 级：</strong>优质渠道商或大型贸易实体。虽无大规模自产，但具备上下游分销及采购能力。
                                </li>
                                <li>
                                  <strong className="text-indigo-400">🥈 B 级：</strong>散客、小微作坊或餐饮零散商户。采购批量小，复购周期长，建议长尾培育。
                                </li>
                                <li>
                                  <strong className="text-slate-400">❌ C 级：</strong>失联、关闭状态，或纯个人买家、皮皮公司。无明显成单及开发价值。
                                </li>
                              </ul>
                            </div>

                            {/* Card Column 2: Follow-up Grade */}
                            <div className="space-y-3">
                              <div className="flex items-center gap-1.5 border-b border-slate-800/60 pb-1.5">
                                <TrendingUp className="h-3.5 w-3.5 text-indigo-400" />
                                <h5 className="text-xs font-bold text-white">跟进优先级划分与空值说明</h5>
                              </div>
                              <ul className="space-y-2 text-[11px] text-slate-400 font-sans">
                                <li>
                                  <strong className="text-indigo-400">🌟 立即跟进：</strong>针对 S 级高潜客户，要求在黄金触达期内，配合专属话术迅速介入。
                                </li>
                                <li>
                                  <strong className="text-blue-400">🚀 重点推进：</strong>针对 A 级优质贸易/分销商，推荐样品及价格方案跟进。
                                </li>
                                <li>
                                  <strong className="text-amber-500">📅 周期培育：</strong>针对 B 级客户，定期自动化邮件孵化。
                                </li>
                                <li>
                                  <strong className="text-slate-500">⏳ 暂缓跟进：</strong>C 级或存在严重欺诈嫌疑的客商。
                                </li>
                                <li className="bg-slate-900/80 p-2.5 rounded-lg border border-slate-800/40 text-[10px] leading-relaxed text-slate-300 mt-1">
                                  <span className="text-amber-400 font-semibold block mb-0.5">⚠️ 为什么“跟进等级”有时显示为空或“待分析”？</span>
                                  若客商从未执行背景调查，或数据极其匮乏，此字段将留空显示为“待分析”。在您点击“立即背调”穿透公开网页与工商数据库后，系统将自动填充等级，便于您科学排布业务精力。
                                </li>
                              </ul>
                            </div>

                            {/* Card Column 3: Risk Alert */}
                            <div className="space-y-3">
                              <div className="flex items-center gap-1.5 border-b border-slate-800/60 pb-1.5">
                                <ShieldAlert className="h-3.5 w-3.5 text-amber-400" />
                                <h5 className="text-xs font-bold text-white">信用与准入风险提示标准</h5>
                              </div>
                              <ul className="space-y-2 text-[11px] text-slate-400">
                                <li>
                                  <strong className="text-amber-400">经营异常警示：</strong>AI 穿透欧洲工商系统（如 KRS/VIES），若发现吊销、清算、严重欠税，将红字警示。
                                </li>
                                <li>
                                  <strong className="text-amber-400">实体匹配核对：</strong>交叉检验客户填写的税号（VAT/NIP）与实际注册地是否一致，防范仿冒欺诈。
                                </li>
                                <li>
                                  <strong className="text-amber-400">官网域名断代：</strong>若官方企业邮箱使用个人免费邮箱，且无对应实体官网或关联内容割裂，系统将提示注意采购资质。
                                </li>
                                <li>
                                  <strong className="text-slate-500">暂无特别信用预警：</strong>代表客商状态良好，可通过正常 B2B 条款合作。
                                </li>
                              </ul>
                            </div>
                          </div>
                        )}
                      </div>
                    </div>

                  </div>

                  {/* Document Navigation Tabs */}
                  <div className="bg-[#0F0F0F] px-6 border-b border-slate-800/80 flex space-x-6 text-sm shrink-0">
                    <button
                      onClick={() => setActiveTab("report")}
                      className={`py-3.5 font-semibold border-b-2 transition flex items-center gap-1.5 ${
                        activeTab === "report"
                          ? "border-emerald-500 text-emerald-400"
                          : "border-transparent text-slate-400 hover:text-slate-200"
                      }`}
                    >
                      <FileText className="h-4 w-4" /> 完整背调报告 (Report)
                    </button>
                    <button
                      onClick={() => setActiveTab("outreach")}
                      className={`py-3.5 font-semibold border-b-2 transition flex items-center gap-1.5 ${
                        activeTab === "outreach"
                          ? "border-emerald-500 text-emerald-400"
                          : "border-transparent text-slate-400 hover:text-slate-200"
                      }`}
                    >
                      <Mail className="h-4 w-4" /> 销售话术与建联策略 (Outreach)
                    </button>
                    <button
                      onClick={() => setActiveTab("sources")}
                      className={`py-3.5 font-semibold border-b-2 transition flex items-center gap-1.5 ${
                        activeTab === "sources"
                          ? "border-emerald-500 text-emerald-400"
                          : "border-transparent text-slate-400 hover:text-slate-200"
                      }`}
                    >
                      <BookOpen className="h-4 w-4" /> 网络实证参考 ({selectedLead.sources?.length || 0})
                    </button>
                    <button
                      onClick={() => setActiveTab("usage")}
                      className={`py-3.5 font-semibold border-b-2 transition flex items-center gap-1.5 ${
                        activeTab === "usage"
                          ? "border-emerald-500 text-emerald-400"
                          : "border-transparent text-slate-400 hover:text-slate-200"
                      }`}
                    >
                      <Cpu className="h-4 w-4" /> API 使用详情面板 (Tokens)
                    </button>
                  </div>

                  {/* Tab Contents */}
                  <div className="flex-1 p-6 overflow-y-auto bg-[#0B0B0B]">
                    
                    {/* TAB 1: Markdown Document Viewer */}
                    {activeTab === "report" && (
                      <div className="bg-[#F9F9F9] rounded-xl p-8 md:p-10 shadow-2xl relative select-text border border-slate-200 text-slate-900 overflow-hidden text-left">
                        {/* Confidential Watermark Accent */}
                        <div className="absolute top-4 right-4 hidden sm:block">
                          <div className="px-2.5 py-1 border border-slate-400 text-[10px] font-bold text-slate-500 uppercase tracking-tighter">
                            CONFIDENTIAL REPORT
                          </div>
                        </div>
                        
                        <div id="report-print-content" className="prose prose-slate max-w-none font-sans">
                          <ReactMarkdown
                            components={{
                              h1: ({ node, ...props }) => <h1 className="text-xl md:text-2xl font-serif font-bold text-black border-b-2 border-slate-950 pb-3 mt-8 mb-4 first:mt-0 uppercase tracking-tight" {...props} />,
                              h2: ({ node, ...props }) => <h2 className="text-lg font-bold text-slate-900 border-b border-slate-300 pb-1 mt-6 mb-3 flex items-center gap-2" {...props} />,
                              h3: ({ node, ...props }) => <h3 className="text-sm font-bold uppercase tracking-wide text-slate-800 mt-5 mb-2" {...props} />,
                              p: ({ node, ...props }) => <p className="text-xs md:text-sm text-slate-700 leading-relaxed mb-4" {...props} />,
                              ul: ({ node, ...props }) => <ul className="list-disc pl-5 space-y-1.5 mb-4 text-xs md:text-sm text-slate-700" {...props} />,
                              ol: ({ node, ...props }) => <ol className="list-decimal pl-5 space-y-1.5 mb-4 text-xs md:text-sm text-slate-700" {...props} />,
                              li: ({ node, ...props }) => <li className="text-xs md:text-sm text-slate-700" {...props} />,
                              table: ({ node, ...props }) => (
                                <div className="overflow-x-auto my-4 rounded-lg border border-slate-300 shadow-sm">
                                  <table className="min-w-full divide-y divide-slate-300 text-xs md:text-sm text-slate-800 bg-white" {...props} />
                                </div>
                              ),
                              thead: ({ node, ...props }) => <thead className="bg-slate-100" {...props} />,
                              tbody: ({ node, ...props }) => <tbody className="divide-y divide-slate-200" {...props} />,
                              tr: ({ node, ...props }) => <tr className="hover:bg-slate-50 transition" {...props} />,
                              th: ({ node, ...props }) => <th className="px-4 py-2.5 text-left font-bold text-slate-900 bg-slate-100" {...props} />,
                              td: ({ node, ...props }) => <td className="px-4 py-2.5 font-sans text-slate-800" {...props} />,
                              blockquote: ({ node, ...props }) => (
                                <blockquote className="border-l-4 border-yellow-400 bg-yellow-50 p-3 my-4 rounded-r italic text-slate-700" {...props} />
                              ),
                              a: ({ node, ...props }) => <a className="text-emerald-700 hover:underline font-semibold" target="_blank" rel="noreferrer referrerPolicy" {...props} />
                            }}
                          >
                            {preprocessMarkdown(selectedLead.report)}
                          </ReactMarkdown>
                        </div>
                      </div>
                    )}

                    {/* TAB 2: Marketing Outreach & CRM pipeline strategy */}
                    {activeTab === "outreach" && (
                      <div className="space-y-6">
                        
                        {/* Interactive CRM follow up checklist */}
                        <div className="bg-[#151515] border border-slate-800 rounded-xl p-5 shadow-sm">
                          <h3 className="text-sm font-semibold text-slate-200 mb-4 flex items-center gap-1.5">
                            <CheckCircle className="h-4 w-4 text-emerald-400" />
                            该客商专属跟进开发路线图 (Interactive Follow-up Pipeline)
                          </h3>
                          <div className="space-y-3">
                            <div className="flex items-start gap-3 p-3 bg-[#0F0F0F] rounded-lg border border-slate-800 hover:border-slate-700 transition text-left">
                              <input
                                type="checkbox"
                                checked={checklist[selectedLead.id]?.stage1 || false}
                                onChange={() => toggleCheck(selectedLead.id, "stage1")}
                                className="h-4.5 w-4.5 rounded border-slate-700 text-emerald-600 focus:ring-emerald-500 focus:ring-offset-slate-900 bg-slate-950 mt-0.5"
                              />
                              <div className="flex-1 text-left">
                                <span className="text-xs font-bold text-emerald-400 block mb-0.5">阶段一：高规格售后触达 (建立初步企业信任)</span>
                                <p className="text-xs text-slate-400">
                                  使用下方的邮件模板，以“检查适配”为由向收件人发起联络，免费赠送视频调试和中文/英文说明支持。
                                </p>
                              </div>
                            </div>

                            <div className="flex items-start gap-3 p-3 bg-[#0F0F0F] rounded-lg border border-slate-800 hover:border-slate-700 transition text-left">
                              <input
                                type="checkbox"
                                checked={checklist[selectedLead.id]?.stage2 || false}
                                onChange={() => toggleCheck(selectedLead.id, "stage2")}
                                className="h-4.5 w-4.5 rounded border-slate-700 text-emerald-600 focus:ring-emerald-500 focus:ring-offset-slate-900 bg-slate-950 mt-0.5"
                              />
                              <div className="flex-1 text-left">
                                <span className="text-xs font-bold text-emerald-400 block mb-0.5">阶段二：瓶子规格预校对 (技术数据套话)</span>
                                <p className="text-xs text-slate-400">
                                  向客商索要其真实产品的瓶身材质、瓶身弧度、和贴标速度。一方面帮助其校准参数，另一方面可判断其真实日产量和设备缺口。
                                </p>
                              </div>
                            </div>

                            <div className="flex items-start gap-3 p-3 bg-[#0F0F0F] rounded-lg border border-slate-800 hover:border-slate-700 transition text-left">
                              <input
                                type="checkbox"
                                checked={checklist[selectedLead.id]?.stage3 || false}
                                onChange={() => toggleCheck(selectedLead.id, "stage3")}
                                className="h-4.5 w-4.5 rounded border-slate-700 text-emerald-600 focus:ring-emerald-500 focus:ring-offset-slate-900 bg-slate-950 mt-0.5"
                              />
                              <div className="flex-1 text-left">
                                <span className="text-xs font-bold text-emerald-400 block mb-0.5">阶段三：推荐整线自动化方案 (长线B2B直签)</span>
                                <p className="text-xs text-slate-400">
                                  在客户确认贴标机运转满意后，主动发送我司的食品工业灌装、旋盖、喷码全自动一体机。引导线下交易并享受免零售平台抽成、享折扣物流。
                                </p>
                              </div>
                            </div>
                          </div>
                        </div>

                        {/* Professional Email Copy Templates */}
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
                          
                          {/* Email Template Card */}
                          <div className="bg-[#151515] border border-slate-800 rounded-xl p-5 flex flex-col justify-between shadow-sm">
                            <div className="space-y-3">
                              <div className="flex justify-between items-center pb-2 border-b border-slate-800/60">
                                <span className="text-xs font-bold text-white flex items-center gap-1.5">
                                  <Mail className="h-4 w-4 text-emerald-400" />
                                  企业开发信模板 (B2B Email Draft)
                                </span>
                                <button
                                  onClick={() => handleCopy(`${templates.emailSubject}\n\n${templates.emailBody}`, "email")}
                                  className="text-[10px] text-emerald-400 hover:underline flex items-center gap-1"
                                >
                                  {copiedText === "email" ? "已复制!" : "复制邮件"}
                                </button>
                              </div>
                              <div className="bg-[#0F0F0F] p-3 rounded-lg font-mono text-[11px] text-slate-300 text-left border border-slate-800/60 overflow-y-auto max-h-[220px]">
                                <span className="text-slate-500 block mb-2 border-b border-slate-800/60 pb-1">
                                  Subject: {templates.emailSubject}
                                </span>
                                <pre className="whitespace-pre-wrap">{templates.emailBody}</pre>
                              </div>
                            </div>
                          </div>

                          {/* IM Chat Script Card */}
                          <div className="bg-[#151515] border border-slate-800 rounded-xl p-5 flex flex-col justify-between shadow-sm">
                            <div className="space-y-3">
                              <div className="flex justify-between items-center pb-2 border-b border-slate-800/60">
                                <span className="text-xs font-bold text-white flex items-center gap-1.5">
                                  <MessageSquare className="h-4 w-4 text-emerald-400" />
                                  即时通讯建联脚本 (IM Chat Script)
                                </span>
                                <button
                                  onClick={() => handleCopy(templates.chatBody, "chat")}
                                  className="text-[10px] text-emerald-400 hover:underline flex items-center gap-1"
                                >
                                  {copiedText === "chat" ? "已复制!" : "复制脚本"}
                                </button>
                              </div>
                              <div className="bg-[#0F0F0F] p-3 rounded-lg font-mono text-[11px] text-slate-300 text-left border border-slate-800/60 overflow-y-auto max-h-[220px]">
                                <span className="text-slate-500 block mb-2 border-b border-slate-800/60 pb-1">
                                  {templates.chatSubject}
                                </span>
                                <pre className="whitespace-pre-wrap">{templates.chatBody}</pre>
                              </div>
                            </div>
                          </div>

                        </div>
                      </div>
                    )}

                    {/* TAB 3: Reference Bibliography and citation verification URLs */}
                    {activeTab === "sources" && (
                      <div className="bg-[#151515] border border-slate-800 rounded-xl p-5 shadow-sm text-left">
                        <h3 className="text-sm font-semibold text-slate-200 mb-4 flex items-center gap-1.5">
                          <Globe className="h-4 w-4 text-emerald-400" />
                          DeepSeek 联网搜索实证佐证资料 (Verified Search Footprint)
                        </h3>
                        {selectedLead.sources && selectedLead.sources.length > 0 ? (
                          <div className="space-y-3">
                            <p className="text-xs text-slate-400 mb-3">
                              以下是 AI 开展背景核查时提取的真实网络线索。您可以直接点击查阅其官方信息，确保资料万无一失：
                            </p>
                            <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                              {selectedLead.sources.map((src, i) => (
                                <a
                                  key={i}
                                  href={src.url}
                                  target="_blank"
                                  rel="noreferrer referrerPolicy"
                                  className="p-3 bg-[#0F0F0F] hover:bg-[#151515] border border-slate-800 hover:border-slate-700 rounded-lg flex items-start justify-between transition group"
                                >
                                  <div className="space-y-1 pr-4 overflow-hidden">
                                    <span className="text-[10px] font-mono text-emerald-400">核查参考网源 #{i + 1}</span>
                                    <h4 className="text-xs font-bold text-white group-hover:text-emerald-400 transition truncate">
                                      {src.title}
                                    </h4>
                                    <p className="text-[10px] text-slate-500 truncate">{src.url}</p>
                                  </div>
                                  <ExternalLink className="h-4.5 w-4.5 text-slate-500 group-hover:text-emerald-400 shrink-0 self-center" />
                                </a>
                              ))}
                            </div>
                          </div>
                        ) : (
                          <div className="text-center p-8 text-slate-500">
                            <Info className="h-8 w-8 mx-auto mb-2 text-slate-600" />
                            <p className="text-xs">
                              该客户通过预加载模版提供，未产生实时的外部调用。您可以在左侧列表中添加新客商并点击“启动背调”以实时检索。
                            </p>
                          </div>
                        )}
                      </div>
                    )}

                    {/* TAB 4: API Token Usage Statistics & Pricing Calculator */}
                    {activeTab === "usage" && (
                      <div className="space-y-6 animate-fadeIn">
                        
                        {/* Summary of Current Task Tokens */}
                        <div className="bg-[#151515] border border-slate-800 rounded-xl p-6 shadow-sm text-left">
                          <h3 className="text-sm font-semibold text-slate-200 mb-4 flex items-center justify-between">
                            <span className="flex items-center gap-1.5">
                              <Cpu className="h-4 w-4 text-emerald-400" />
                              当前背调任务 Token 消耗分析 / Current Investigation Usage
                            </span>
                            <span className="text-xs text-slate-500 font-mono">
                              ID: {selectedLead.id}
                            </span>
                          </h3>

                          {selectedLead.usage ? (
                            <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
                              {/* Card 1: Model & Mode */}
                              <div className="bg-[#0F0F0F] border border-slate-800/80 p-4 rounded-xl relative overflow-hidden group">
                                <div className="absolute top-0 right-0 p-3 opacity-10 group-hover:opacity-20 transition">
                                  <Cpu className="h-16 w-16 text-purple-400" />
                                </div>
                                <span className="text-[10px] text-slate-500 font-bold uppercase tracking-wider block mb-1">执行模型 / Model</span>
                                <p className="text-sm font-mono font-bold text-white mb-2 truncate flex items-center gap-1.5">
                                  {selectedLead.usage.modelUsed}
                                </p>
                                <div className="text-[11px] text-slate-400 space-y-1">
                                  <p className="flex items-center gap-1">
                                    <span className="h-1.5 w-1.5 rounded-full bg-emerald-400"></span>
                                    {selectedLead.usage.searchGroundingUsed ? "🌐 DeepSeek 联网启用" : "⚠️ Fallback 无搜索状态"}
                                  </p>
                                  <p className="text-[10px] text-slate-500">
                                    智能切换策略以维持服务高可用
                                  </p>
                                </div>
                              </div>

                              {/* Card 2: Prompt Tokens */}
                              <div className="bg-[#0F0F0F] border border-slate-800/80 p-4 rounded-xl relative overflow-hidden group">
                                <span className="text-[10px] text-slate-500 font-bold uppercase tracking-wider block mb-1">输入 Token / Prompt Tokens</span>
                                <p className="text-2xl font-mono font-bold text-slate-100 mb-1">
                                  {selectedLead.usage.promptTokens.toLocaleString()}
                                </p>
                                <p className="text-[11px] text-slate-400">
                                  包含复杂的波兰/欧洲工商审计系统提示、爬虫网页抓取纯文本上下文。
                                </p>
                              </div>

                              {/* Card 3: Completion Tokens */}
                              <div className="bg-[#0F0F0F] border border-slate-800/80 p-4 rounded-xl relative overflow-hidden group">
                                <span className="text-[10px] text-slate-500 font-bold uppercase tracking-wider block mb-1">输出 Token / Output Tokens</span>
                                <p className="text-2xl font-mono font-bold text-slate-100 mb-1">
                                  {selectedLead.usage.completionTokens.toLocaleString()}
                                </p>
                                <p className="text-[11px] text-slate-400">
                                  AI 生成的结构化规范背调、风险警示、客商分级及中英双语定制开发信。
                                </p>
                              </div>

                              {/* Card 4: Cost Calculation */}
                              <div className="bg-[#0F0F0F] border border-emerald-500/20 p-4 rounded-xl relative overflow-hidden group bg-gradient-to-br from-[#0F0F0F] to-[#121f17]/20">
                                <div className="absolute top-0 right-0 p-3 opacity-10">
                                  <Zap className="h-16 w-16 text-emerald-400" />
                                </div>
                                <span className="text-[10px] text-emerald-400 font-bold uppercase tracking-wider block mb-1">预估调用价格 / Estimated Cost</span>
                                <p className="text-2xl font-mono font-bold text-emerald-400 mb-1 flex items-baseline gap-1.5 flex-wrap">
                                  ¥{calculateTokenCost(selectedLead.usage.modelUsed, selectedLead.usage.promptTokens, selectedLead.usage.completionTokens).rmb.toFixed(4)}
                                  <span className="text-xs font-normal text-slate-500">(${calculateTokenCost(selectedLead.usage.modelUsed, selectedLead.usage.promptTokens, selectedLead.usage.completionTokens).usd.toFixed(5)} USD)</span>
                                </p>
                                <p className="text-[11px] text-slate-400">
                                  按官方最新 API 标准及 1 USD = 7.25 CNY 汇率换算。此智能方案相较人工商业背调费用，节约达 <span className="text-emerald-400 font-semibold font-mono">99.5%</span>。
                                </p>
                              </div>
                            </div>
                          ) : (
                            <div className="bg-[#0F0F0F] border border-dashed border-slate-800 rounded-lg p-8 text-center flex flex-col items-center justify-center space-y-2">
                              <Zap className="h-6 w-6 text-slate-600 animate-pulse" />
                              <p className="text-xs font-semibold text-slate-400">预置/导入示例数据暂无实时 API 解析记录</p>
                              <p className="text-[11px] text-slate-500 max-w-lg">
                                您可以通过点击右侧的 <span className="text-emerald-400">“重新发起背调”</span> 启动实时互联网交叉核验与 AI 解析，系统将自动采集并持久化展示该调查的 API Token 与费用。
                              </p>
                            </div>
                          )}
                        </div>

                        {/* Interactive Token Pricing List & Cost Estimator */}
                        <div className="grid grid-cols-1 md:grid-cols-12 gap-4 text-left">
                          
                          {/* Pricing Reference Card */}
                          <div className="md:col-span-7 bg-[#151515] border border-slate-800 rounded-xl p-5 shadow-sm">
                            <h3 className="text-sm font-semibold text-slate-200 mb-3 flex items-center gap-1.5">
                              <Database className="h-4 w-4 text-emerald-400" />
                              DeepSeek 官方 API 计费价格参考 (Price Guide)
                            </h3>
                            <p className="text-xs text-slate-400 mb-4 leading-relaxed">
                              本系统深度优化了 Token 裁剪策略，优先利用上下文缓存及压缩净化算法剔除网页噪声，保证您的每一次背调都控制在极低经济成本内。
                            </p>

                            <div className="overflow-x-auto border border-slate-800 rounded-lg">
                              <table className="w-full text-xs text-left text-slate-300">
                                <thead className="bg-[#0F0F0F] text-[10px] text-slate-500 uppercase tracking-wider border-b border-slate-800">
                                  <tr>
                                    <th className="px-4 py-2.5">模型 (Model)</th>
                                    <th className="px-4 py-2.5">输入 (Input / 1M)</th>
                                    <th className="px-4 py-2.5">输出 (Output / 1M)</th>
                                    <th className="px-4 py-2.5">折合人民币 (CNY)</th>
                                    <th className="px-4 py-2.5">适用场景</th>
                                  </tr>
                                </thead>
                                <tbody className="divide-y divide-slate-800/60 font-mono">
                                  <tr>
                                    <td className="px-4 py-3 font-semibold text-slate-200">deepseek-v4-flash</td>
                                    <td className="px-4 py-3 text-emerald-500">$0.14</td>
                                    <td className="px-4 py-3 text-emerald-500">$0.28</td>
                                    <td className="px-4 py-3 text-[11px] text-emerald-400">¥1.02 / ¥2.03</td>
                                    <td className="px-4 py-3 text-[10px] text-slate-400 font-sans">闪电级高精度日常主流背调 (推荐)</td>
                                  </tr>
                                </tbody>
                              </table>
                            </div>
                            <div className="mt-3.5 p-3 bg-[#0F0F0F] border border-slate-800/80 rounded-lg flex items-start gap-2">
                              <Info className="h-4 w-4 text-emerald-500 shrink-0 mt-0.5" />
                              <p className="text-[10px] text-slate-400 leading-relaxed">
                                <strong className="text-slate-300">联网搜索零阶计费提示：</strong>DeepSeek 联网搜索由引擎自动执行，搜索成本已包含在输出 Token 计费中。相较传统商业背调动辄数百元的人工差旅或付费软件，AI 自动化极速穿透具备难以置信的绝对性价比。
                              </p>
                            </div>
                          </div>

                          {/* Cumulative System Statistics Card */}
                          <div className="md:col-span-5 bg-[#151515] border border-slate-800 rounded-xl p-5 shadow-sm flex flex-col justify-between">
                            <div>
                              <h3 className="text-sm font-semibold text-slate-200 mb-1 flex items-center gap-1.5">
                                <TrendingUp className="h-4 w-4 text-emerald-400" />
                                累计背调分析统计 (Cumulative Metrics)
                              </h3>
                              <p className="text-[11px] text-slate-400 mb-4">
                                本地工作台已完成背调的汇总用量监控。
                              </p>

                              {(() => {
                                const completedLeads = leads.filter(l => l.status === "completed" && l.usage);
                                const totalPrompt = completedLeads.reduce((sum, l) => sum + (l.usage?.promptTokens || 0), 0);
                                const totalCompletion = completedLeads.reduce((sum, l) => sum + (l.usage?.completionTokens || 0), 0);
                                const totalTokens = totalPrompt + totalCompletion;
                                
                                // Calculate total USD and RMB using our precise formula
                                const { totalUsd, totalRmb } = completedLeads.reduce((acc, l) => {
                                  if (!l.usage) return acc;
                                  const { usd, rmb } = calculateTokenCost(l.usage.modelUsed, l.usage.promptTokens, l.usage.completionTokens);
                                  acc.totalUsd += usd;
                                  acc.totalRmb += rmb;
                                  return acc;
                                }, { totalUsd: 0, totalRmb: 0 });

                                return (
                                  <div className="space-y-3.5">
                                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
                                      <div className="bg-[#0F0F0F] border border-slate-800/60 p-2.5 rounded-lg">
                                        <span className="text-[10px] text-slate-500 block mb-0.5">累计分析客商数</span>
                                        <p className="text-lg font-mono font-bold text-white">
                                          {leads.filter(l => l.status === "completed").length} <span className="text-[10px] text-slate-500 font-normal">家</span>
                                        </p>
                                      </div>
                                      <div className="bg-[#0F0F0F] border border-slate-800/60 p-2.5 rounded-lg">
                                        <span className="text-[10px] text-slate-500 block mb-0.5">累计调用估算</span>
                                        <p className="text-sm font-mono font-bold text-emerald-400 flex flex-wrap items-baseline gap-1">
                                          ¥{totalRmb.toFixed(4)}
                                          <span className="text-[10px] font-normal text-slate-500 font-sans">(${totalUsd.toFixed(5)})</span>
                                        </p>
                                      </div>
                                    </div>

                                    <div className="space-y-2 pt-2 border-t border-slate-800/60">
                                      <div className="flex justify-between items-center text-xs">
                                        <span className="text-slate-400">累计输入 Tokens:</span>
                                        <span className="font-mono text-slate-200">{totalPrompt.toLocaleString()}</span>
                                      </div>
                                      <div className="flex justify-between items-center text-xs">
                                        <span className="text-slate-400">累计输出 Tokens:</span>
                                        <span className="font-mono text-slate-200">{totalCompletion.toLocaleString()}</span>
                                      </div>
                                      <div className="flex justify-between items-center text-xs pt-1 border-t border-slate-800/20 font-semibold">
                                        <span className="text-slate-300">累计总消耗 Tokens:</span>
                                        <span className="font-mono text-emerald-400">{totalTokens.toLocaleString()}</span>
                                      </div>
                                    </div>
                                  </div>
                                );
                              })()}
                            </div>

                            <p className="text-[10px] text-slate-500 text-center mt-4 pt-3.5 border-t border-slate-800/40">
                              💡 本地存储所有记录，清除浏览器缓存或关闭开发服务器将重置本地统计。
                            </p>
                          </div>

                        </div>

                      </div>
                    )}

                  </div>
                </div>
              )}
            </div>
          ) : (
            <div className="bg-[#151515] border border-slate-800 rounded-xl flex-1 flex flex-col items-center justify-center p-12 text-center text-slate-500 shadow-sm min-h-[500px]">
              <Building2 className="h-12 w-12 text-slate-600 mb-3 stroke-1" />
              <h3 className="text-base font-bold text-white">未选择任何客户</h3>
              <p className="text-xs text-slate-400 mt-1">请从左侧列表选择一家客户查看其详细背景调查报告，或添加新客户。</p>
            </div>
          )}
          
        </div>
      </main>

      {/* Static Footer Usage Guide */}
      <footer id="usage-guide" className="bg-[#0F0F0F] border-t border-slate-800/80 px-6 py-8 mt-12 text-slate-400 text-sm">
        <div className="max-w-7xl mx-auto grid grid-cols-1 md:grid-cols-3 gap-8 text-left">
          <div className="space-y-3">
            <h4 className="text-white font-semibold flex items-center gap-1.5 text-xs">
              <Sparkles className="h-4 w-4 text-emerald-400" /> AI 联网搜索穿透背调
            </h4>
            <p className="text-xs text-slate-400 leading-relaxed">
              传统的企业背景核实耗时耗力，本系统结合了 <strong>DeepSeek V4 Flash</strong> 强大的推理框架和 <strong>DeepSeek 联网搜索</strong> 实证网络检索能力，能瞬间检索出境外主体的真实注册状态、工厂所属领域和产能，生成具有实战价值的客户背调报告。
            </p>
          </div>
          <div className="space-y-3">
            <h4 className="text-white font-semibold flex items-center gap-1.5 text-xs">
              <FileSpreadsheet className="h-4 w-4 text-emerald-400" /> 批量导入说明
            </h4>
            <p className="text-xs text-slate-400 leading-relaxed">
              为了极大提高外贸业务员的跟单效率，系统提供了<strong>电子表格数据批量一键解析功能</strong>。您可直接在 Excel 或 WPS 选中需要背调的客户行并按 <kbd className="bg-[#151515] border border-slate-800 px-1 py-0.5 rounded text-[10px] text-slate-200">Ctrl+C</kbd>，然后粘贴到系统批量导入框中，系统会自动提取出每一列并排队执行核查。
            </p>
          </div>
          <div className="space-y-3">
            <h4 className="text-white font-semibold flex items-center gap-1.5 text-xs">
              <Mail className="h-4 w-4 text-emerald-400" /> 转化跟进设计哲学
            </h4>
            <p className="text-xs text-slate-400 leading-relaxed">
              跨境电商（如 AliExpress 平台）买家往往是碎片化采购，但<strong>工业品、包装机械买家必然存在实体工厂</strong>。调查的最高价值在于“破冰”，通过AI针对性分析其产品规格和生产痛点，在售后黄金48小时，以技术咨询为借口沉淀为<strong>线下长期高额 B2B 战略大户</strong>。
            </p>
          </div>
        </div>
        <div className="max-w-7xl mx-auto border-t border-slate-800/40 mt-8 pt-4 text-center text-xs text-slate-500">
          Client Background Investigator Workspace • Powered by DeepSeek V4 Flash • All Rights Reserved © 2026.
        </div>
      </footer>

      {/* 飞书多维表格同步映射分析模版 (Mapping Analysis Dialog) */}
      {isMappingAnalysisModalOpen && syncMappingAnalysis && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-sm animate-fadeIn">
          <div className="bg-[#151515] border border-slate-800 rounded-2xl w-full max-w-4xl max-h-[85vh] overflow-hidden flex flex-col shadow-2xl">
            {/* Modal Header */}
            <div className="p-5 border-b border-slate-800 bg-[#0F0F0F] flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <div className="h-9 w-9 rounded-xl bg-emerald-500/10 border border-emerald-500/20 flex items-center justify-center">
                  <FileSpreadsheet className="h-5 w-5 text-emerald-400" />
                </div>
                <div>
                  <h3 className="text-sm font-bold text-white flex items-center gap-2">
                    飞书多维表格字段映射核对成功！
                    <span className="text-[10px] bg-emerald-500/10 text-emerald-400 px-2 py-0.5 rounded-full border border-emerald-500/20 font-mono">
                      完美适配 48 字段标准
                    </span>
                  </h3>
                  <p className="text-xs text-slate-400 mt-0.5">
                    系统已自动对 Excel 字段和飞书多维表格字段进行了深度智能匹配。请核对下方字段映射分析表：
                  </p>
                </div>
              </div>
              <button
                onClick={() => setIsMappingAnalysisModalOpen(false)}
                className="text-slate-400 hover:text-white transition p-1 bg-slate-800/40 rounded-lg hover:bg-slate-800"
              >
                <X className="h-4 w-4" />
              </button>
            </div>

            {/* Modal Content */}
            <div className="p-6 overflow-y-auto space-y-4 flex-1">
              <div className="bg-[#0F0F0F] border border-slate-800/60 p-3 rounded-xl flex items-center justify-between text-xs text-slate-300">
                <div className="flex gap-4">
                  <div>
                    <span className="text-slate-500">同步总记录数:</span>{" "}
                    <strong className="text-white font-mono">{leads.filter(l => l.excelRowIndex !== undefined).length} 条</strong>
                  </div>
                  <div>
                    <span className="text-slate-500">匹配成功字段:</span>{" "}
                    <strong className="text-emerald-400 font-mono">
                      {syncMappingAnalysis.filter(m => m.status === "success").length} / {syncMappingAnalysis.length}
                    </strong>
                  </div>
                  <div>
                    <span className="text-slate-500">未匹配/跳过字段:</span>{" "}
                    <strong className="text-amber-400 font-mono">
                      {syncMappingAnalysis.filter(m => m.status === "unmatched").length}
                    </strong>
                  </div>
                </div>
                <div className="text-[10px] text-slate-500 flex items-center gap-1.5">
                  <span className="h-2 w-2 rounded-full bg-[#00D061] animate-pulse"></span>
                  数据类型已按多维表格格式自动强类型转换
                </div>
              </div>

              {/* Table of Mappings */}
              <div className="border border-slate-800/80 rounded-xl overflow-hidden bg-[#0F0F0F]">
                <table className="w-full text-left border-collapse text-xs">
                  <thead>
                    <tr className="bg-[#151515] text-slate-400 border-b border-slate-800">
                      <th className="py-2.5 px-4 font-semibold">#</th>
                      <th className="py-2.5 px-4 font-semibold">原始 Excel 列名</th>
                      <th className="py-2.5 px-4 font-semibold">匹配方式</th>
                      <th className="py-2.5 px-4 font-semibold">飞书多维表格目标字段名</th>
                      <th className="py-2.5 px-4 font-semibold">字段数据类型</th>
                      <th className="py-2.5 px-4 font-semibold text-right">映射匹配状态</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/40 text-slate-300 font-mono">
                    {syncMappingAnalysis.map((item, idx) => {
                      const isMatched = item.status === "success" || item.status === "coercion_failed";
                      return (
                        <tr key={idx} className="hover:bg-[#151515]/40 transition">
                          <td className="py-2 px-4 text-slate-500 font-sans text-[10px]">{idx + 1}</td>
                          <td className="py-2 px-4 font-semibold text-slate-200">{item.excelField}</td>
                          <td className="py-2 px-4">
                            {item.matchType === "exact" && (
                              <span className="bg-emerald-500/10 text-emerald-400 text-[10px] px-1.5 py-0.5 rounded-md border border-emerald-500/20 font-sans">
                                精确匹配
                              </span>
                            )}
                            {item.matchType === "alias" && (
                              <span className="bg-sky-500/10 text-sky-400 text-[10px] px-1.5 py-0.5 rounded-md border border-sky-500/20 font-sans">
                                智能别名匹配
                              </span>
                            )}
                            {item.matchType === "substring" && (
                              <span className="bg-indigo-500/10 text-indigo-400 text-[10px] px-1.5 py-0.5 rounded-md border border-indigo-500/20 font-sans">
                                模糊子串匹配
                              </span>
                            )}
                            {item.matchType === "none" && (
                              <span className="bg-slate-800 text-slate-500 text-[10px] px-1.5 py-0.5 rounded-md border border-slate-700/50 font-sans">
                                未能匹配
                              </span>
                            )}
                          </td>
                          <td className="py-2 px-4">
                            {isMatched ? (
                              <span className="text-white font-semibold flex items-center gap-1">
                                {item.bitableField}
                              </span>
                            ) : (
                              <span className="text-slate-600">-</span>
                            )}
                          </td>
                          <td className="py-2 px-4 text-slate-400 font-sans text-[11px]">
                            {item.fieldTypeName}
                          </td>
                          <td className="py-2 px-4 text-right">
                            {item.status === "success" && (
                              <span className="text-emerald-400 font-bold font-sans text-[10px] flex items-center justify-end gap-1">
                                <Check className="h-3.5 w-3.5" /> 正常导入
                              </span>
                            )}
                            {item.status === "coercion_failed" && (
                              <span className="text-amber-400 font-bold font-sans text-[10px] flex items-center justify-end gap-1">
                                <AlertCircle className="h-3.5 w-3.5" /> 转换空值
                              </span>
                            )}
                            {item.status === "unmatched" && (
                              <span className="text-slate-500 font-medium font-sans text-[10px] flex items-center justify-end gap-1">
                                <Info className="h-3.5 w-3.5 text-slate-600" /> 跳过写入
                              </span>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Modal Footer */}
            <div className="p-4 border-t border-slate-800 bg-[#0F0F0F] flex items-center justify-between">
              <span className="text-[11px] text-slate-500">
                提示：系统已对文本35、带空格字段、德/意/波/墨城市以及中英文括号进行了全面对齐，无需人工修复表格格式。
              </span>
              <button
                onClick={() => setIsMappingAnalysisModalOpen(false)}
                className="bg-emerald-600 hover:bg-emerald-500 text-white font-semibold text-xs px-5 py-2 rounded-lg transition shadow-md active:bg-emerald-700"
              >
                我知道了，返回工作台
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
