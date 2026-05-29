const fs = require("fs");
const path = require("path");
const { spawn } = require("child_process");
const {
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  TextRun,
} = require("docx");

const ROOT = __dirname;
const activeCodexMessages = new Set();
const latestOfficeFiles = new Map();
const latestImages = new Map();

function loadEnvFile(filePath) {
  if (!fs.existsSync(filePath)) return;
  const lines = fs.readFileSync(filePath, "utf8").split(/\r?\n/);
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const index = trimmed.indexOf("=");
    if (index < 0) continue;
    const key = trimmed.slice(0, index).trim();
    let value = trimmed.slice(index + 1).trim();
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }
    if (!process.env[key]) process.env[key] = value;
  }
}

function requireEnv(name) {
  const value = process.env[name];
  if (!value) throw new Error(`Missing required config: ${name}`);
  return value;
}

function parseAllowedOpenIds(value) {
  return new Set(
    String(value || "")
      .split(",")
      .map((item) => item.trim())
      .filter(Boolean)
  );
}

function getEventPayload(data) {
  return data && data.event ? data.event : data;
}

function getTextFromContent(content) {
  if (!content) return "";
  try {
    const parsed = typeof content === "string" ? JSON.parse(content) : content;
    return parsed.text || parsed.content || "";
  } catch {
    return String(content);
  }
}

function parseMessageContent(content) {
  if (!content) return {};
  try {
    return typeof content === "string" ? JSON.parse(content) : content;
  } catch {
    return {};
  }
}

function getSenderOpenId(payload) {
  return (
    payload?.sender?.sender_id?.open_id ||
    payload?.sender?.sender_id?.user_id ||
    payload?.sender?.open_id ||
    ""
  );
}

function getConversationKey(chatId, senderOpenId) {
  return `${chatId || "chat"}:${senderOpenId || "sender"}`;
}

function isFromBot(payload) {
  return payload?.sender?.sender_type === "app";
}

function stripCodexBridgePrefix(text) {
  return text.trim().replace(/^(codex|问codex|让codex|@codex|ai|问ai)[：:\s]+/i, "").trim();
}

function splitText(text, maxLength = 1500) {
  const chunks = [];
  let rest = String(text || "");
  while (rest.length > maxLength) {
    let index = rest.lastIndexOf("\n", maxLength);
    if (index < maxLength * 0.5) index = maxLength;
    chunks.push(rest.slice(0, index));
    rest = rest.slice(index).trimStart();
  }
  if (rest) chunks.push(rest);
  return chunks;
}

function rmOldPath(targetPath, cutoffMs, stats) {
  if (!fs.existsSync(targetPath)) return;
  const stat = fs.statSync(targetPath);
  if (stat.isDirectory()) {
    for (const entry of fs.readdirSync(targetPath)) {
      rmOldPath(path.join(targetPath, entry), cutoffMs, stats);
    }
    const remaining = fs.readdirSync(targetPath);
    const refreshed = fs.statSync(targetPath);
    if (remaining.length === 0 && refreshed.mtimeMs < cutoffMs) {
      fs.rmSync(targetPath, { recursive: true, force: true });
      stats.dirs += 1;
    }
    return;
  }
  if (stat.mtimeMs < cutoffMs) {
    stats.bytes += stat.size;
    fs.rmSync(targetPath, { force: true });
    stats.files += 1;
  }
}

function cleanupOldFiles(retentionDays) {
  const days = Number(retentionDays || 7);
  if (!Number.isFinite(days) || days <= 0) return { files: 0, dirs: 0, bytes: 0 };
  const cutoffMs = Date.now() - days * 24 * 60 * 60 * 1000;
  const stats = { files: 0, dirs: 0, bytes: 0 };
  const targets = [
    path.join(ROOT, "uploads"),
    path.join(ROOT, "outputs", "jobs"),
  ];

  for (const target of targets) {
    rmOldPath(target, cutoffMs, stats);
  }

  const outputDir = path.join(ROOT, "outputs");
  if (fs.existsSync(outputDir)) {
    for (const entry of fs.readdirSync(outputDir)) {
      if (/^codex-.*\.(log|txt)$/i.test(entry)) {
        rmOldPath(path.join(outputDir, entry), cutoffMs, stats);
      }
    }
  }
  return stats;
}

function scheduleCleanup(retentionDays) {
  const run = () => {
    try {
      const stats = cleanupOldFiles(retentionDays);
      console.log(
        `[cleanup] removed ${stats.files} files, ${stats.dirs} dirs, ${Math.round(stats.bytes / 1024)} KB`
      );
    } catch (error) {
      console.error("[cleanup] failed", error);
    }
  };
  run();
  setInterval(run, 24 * 60 * 60 * 1000);
}

function wantsWordFile(prompt) {
  const wordTerms = ["\u6587\u6863", "\u62a5\u544a", "\u7b80\u62a5", "\u6750\u6599", "\u65b9\u6848\u4e66", "\u8bf4\u660e\u4e66", "\u7a3f\u4ef6", "\u6587\u4ef6"];
  if (wordTerms.some((term) => String(prompt).includes(term))) return true;
  return /(word|docx|文档|报告|简报|材料|方案书|说明书|稿件|文件)/i.test(prompt);
}

function buildWordFilePrompt(userPrompt) {
  return [
    "You are running inside a local Codex Bridge on the user's home computer.",
    "The user wants a Word document file, not just pasted text.",
    "Create a valid .docx file in the current workspace under the outputs directory.",
    "Use the installed Node.js docx package if useful. You may create a short helper script and run it.",
    "Do not say that the workspace is read-only. This task is running with workspace-write permission.",
    "Write the document content in the same language as the user's request unless the user asks otherwise.",
    "Make the document polished and practical, with a clear title, headings, paragraphs, and bullet points where appropriate.",
    "When finished, reply only with a short summary and the relative file path of the created .docx file.",
    "",
    `User request: ${userPrompt}`,
  ].join("\n");
}

function findNewDocxFiles(sinceMs) {
  const outputDir = path.join(ROOT, "outputs");
  if (!fs.existsSync(outputDir)) return [];
  return fs
    .readdirSync(outputDir)
    .filter((name) => name.toLowerCase().endsWith(".docx"))
    .map((name) => {
      const filePath = path.join(outputDir, name);
      const stat = fs.statSync(filePath);
      return { filePath, mtimeMs: stat.mtimeMs };
    })
    .filter((item) => item.mtimeMs >= sinceMs - 2000)
    .sort((a, b) => b.mtimeMs - a.mtimeMs)
    .map((item) => item.filePath);
}

function wantsPptFile(prompt) {
  const pptTerms = ["\u5e7b\u706f\u7247", "\u6f14\u793a", "\u6c47\u62a5", "\u8def\u6f14", "\u8bfe\u4ef6"];
  if (pptTerms.some((term) => String(prompt).includes(term))) return true;
  return /(ppt|pptx|powerpoint)/i.test(prompt);
}

function getOfficeFileType(prompt) {
  if (wantsPptFile(prompt)) return "pptx";
  if (wantsWordFile(prompt)) return "docx";
  return "";
}

function getOfficeFileTypeByName(fileName) {
  const ext = path.extname(fileName || "").toLowerCase();
  if (ext === ".ppt" || ext === ".pptx") return "pptx";
  if (ext === ".doc" || ext === ".docx") return "docx";
  return "";
}

function wantsEditExistingFile(prompt) {
  const terms = [
    "\u4fee\u6539",
    "\u6539\u4e00\u4e0b",
    "\u6539\u6210",
    "\u6539\u4e3a",
    "\u6362\u6210",
    "\u6362\u4e3a",
    "\u53d8\u6210",
    "\u8c03\u6574",
    "\u66ff\u6362",
    "\u66f4\u65b0",
    "\u6da6\u8272",
    "\u4f18\u5316",
    "\u8865\u5145",
    "\u7b2c",
    "\u9875",
    "\u6bb5",
    "\u8fd9\u4efd",
    "\u521a\u624d",
    "\u4e0a\u4e00\u4efd",
    "\u6587\u4ef6",
  ];
  return terms.some((term) => String(prompt).includes(term));
}

function wantsImageTask(prompt) {
  const terms = [
    "\u56fe\u7247",
    "\u7167\u7247",
    "\u62cd\u7167",
    "\u8fd9\u5f20\u56fe",
    "\u56fe\u91cc",
    "\u56fe\u4e2d",
    "\u8bc6\u522b",
    "\u770b\u770b",
    "\u5206\u6790",
    "\u63d0\u53d6",
    "\u626b\u63cf",
    "\u8f6c\u6587\u5b57",
  ];
  if (terms.some((term) => String(prompt).includes(term))) return true;
  return /(image|photo|picture|screenshot|ocr)/i.test(prompt);
}

function createOfficeJobDir() {
  const dir = path.join(ROOT, "outputs", "jobs", String(Date.now()));
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

function buildOfficeFilePrompt(userPrompt, fileType) {
  const appName = fileType === "pptx" ? "Microsoft PowerPoint" : "Microsoft Word";
  const suffix = fileType === "pptx" ? ".pptx" : ".docx";
  return [
    "You are running inside a local Codex Bridge on the user's home computer.",
    `The user wants a real ${suffix} file, not just pasted text.`,
    `Create a valid ${suffix} file in the current working directory.`,
    `Use ${appName} automation from PowerShell/COM if available. You may create and run a short helper script.`,
    "Prefer the actual Office application automation path over wrapping plain text yourself.",
    "Do not say that the workspace is read-only. This task is running with workspace-write permission.",
    "Safety rule: only create a new file for this task. Do not delete, rename, overwrite, or modify any existing file.",
    "Do not write outside the current working directory.",
    "Write the content in the same language as the user's request unless the user asks otherwise.",
    fileType === "pptx"
      ? "Make the presentation polished and practical, with a clear title slide, concise page titles, and useful bullet points."
      : "Make the document polished and practical, with a clear title, headings, paragraphs, and bullet points where appropriate.",
    `When finished, reply only with a short summary and the relative file path of the created ${suffix} file.`,
    "",
    `User request: ${userPrompt}`,
  ].join("\n");
}

function buildEditOfficeFilePrompt(userPrompt, inputFileName, fileType) {
  const suffix = fileType === "pptx" ? ".pptx" : ".docx";
  const appName = fileType === "pptx" ? "Microsoft PowerPoint" : "Microsoft Word";
  return [
    "You are running inside a local Codex Bridge on the user's home computer.",
    `The user uploaded an existing ${suffix} file and wants a modified copy.`,
    `Input file in the current working directory: ${inputFileName}`,
    `Use ${appName} automation from PowerShell/COM if available.`,
    "Safety rule: do not delete, rename, or move any file.",
    "Do not modify files outside the current working directory.",
    "Create a new modified output file in the current working directory.",
    `The modified output file must be a valid ${suffix} file.`,
    "You may edit the local copy of the uploaded file if needed, but do not delete it.",
    "When finished, reply only with a short summary and the relative file path of the modified file.",
    "",
    `User modification request: ${userPrompt}`,
  ].join("\n");
}

function findNewOfficeFiles(dir, sinceMs, fileType, excludeNames = []) {
  if (!fs.existsSync(dir)) return [];
  const ext = `.${fileType}`;
  const exclude = new Set(excludeNames.map((name) => String(name).toLowerCase()));
  const files = fs
    .readdirSync(dir)
    .filter((name) => name.toLowerCase().endsWith(ext))
    .filter((name) => !exclude.has(name.toLowerCase()))
    .map((name) => {
      const filePath = path.join(dir, name);
      const stat = fs.statSync(filePath);
      return { filePath, mtimeMs: stat.mtimeMs };
    })
    .sort((a, b) => b.mtimeMs - a.mtimeMs);
  const fresh = files.filter((item) => item.mtimeMs >= sinceMs - 2000);
  return (fresh.length ? fresh : files).map((item) => item.filePath);
}

function safeUploadedName(fileName, fallback = "uploaded-file") {
  const ext = path.extname(fileName || "");
  const base = path.basename(fileName || fallback, ext).replace(/[\\/:*?"<>|]/g, "").slice(0, 60);
  return `${base || fallback}${ext || ""}`;
}

function findLatestUploadedOfficeFile() {
  const uploadRoot = path.join(ROOT, "uploads");
  if (!fs.existsSync(uploadRoot)) return null;
  const found = [];
  const visit = (dir) => {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      const fullPath = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        visit(fullPath);
        continue;
      }
      const fileType = getOfficeFileTypeByName(entry.name);
      if (!fileType) continue;
      const stat = fs.statSync(fullPath);
      found.push({
        filePath: fullPath,
        fileName: entry.name,
        fileType,
        receivedAt: stat.mtimeMs,
      });
    }
  };
  visit(uploadRoot);
  found.sort((a, b) => b.receivedAt - a.receivedAt);
  return found[0] || null;
}

function getImageExtension(content) {
  const imageName = content.file_name || content.fileName || content.name || "";
  const ext = path.extname(imageName).toLowerCase();
  if ([".jpg", ".jpeg", ".png", ".webp", ".gif"].includes(ext)) return ext;
  return ".jpg";
}

function detectImageExtension(filePath, fallback = ".jpg") {
  const buffer = fs.readFileSync(filePath);
  if (buffer.length >= 3 && buffer[0] === 0xff && buffer[1] === 0xd8 && buffer[2] === 0xff) {
    return ".jpg";
  }
  if (
    buffer.length >= 8 &&
    buffer[0] === 0x89 &&
    buffer[1] === 0x50 &&
    buffer[2] === 0x4e &&
    buffer[3] === 0x47
  ) {
    return ".png";
  }
  if (
    buffer.length >= 12 &&
    buffer.slice(0, 4).toString("ascii") === "RIFF" &&
    buffer.slice(8, 12).toString("ascii") === "WEBP"
  ) {
    return ".webp";
  }
  if (buffer.length >= 6 && buffer.slice(0, 3).toString("ascii") === "GIF") {
    return ".gif";
  }
  return fallback;
}

function findLatestUploadedImage() {
  const imageRoot = path.join(ROOT, "uploads", "images");
  if (!fs.existsSync(imageRoot)) return null;
  const found = [];
  const visit = (dir) => {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      const fullPath = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        visit(fullPath);
        continue;
      }
      if (!/\.(jpg|jpeg|png|webp|gif)$/i.test(entry.name)) continue;
      const stat = fs.statSync(fullPath);
      found.push({
        filePath: fullPath,
        fileName: entry.name,
        receivedAt: stat.mtimeMs,
      });
    }
  };
  visit(imageRoot);
  found.sort((a, b) => b.receivedAt - a.receivedAt);
  return found[0] || null;
}

function rememberLatestOfficeFile(chatId, senderOpenId, filePath, fileType) {
  latestOfficeFiles.set(getConversationKey(chatId, senderOpenId), {
    filePath,
    fileName: path.basename(filePath),
    fileType,
    receivedAt: Date.now(),
  });
}

function rememberLatestImage(chatId, senderOpenId, filePath) {
  latestImages.set(getConversationKey(chatId, senderOpenId), {
    filePath,
    fileName: path.basename(filePath),
    receivedAt: Date.now(),
  });
}

function safeFilePart(value) {
  const cleaned = String(value || "")
    .replace(/[\\/:*?"<>|]/g, "")
    .replace(/\s+/g, "-")
    .slice(0, 36);
  return cleaned || "codex-document";
}

function paragraphFromLine(line) {
  const text = line.trim();
  if (!text) {
    return new Paragraph({ text: "" });
  }
  if (/^#{1,3}\s+/.test(text)) {
    return new Paragraph({
      text: text.replace(/^#{1,3}\s+/, ""),
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 240, after: 120 },
    });
  }
  if (/^[-*]\s+/.test(text)) {
    return new Paragraph({
      text: text.replace(/^[-*]\s+/, ""),
      bullet: { level: 0 },
      spacing: { after: 80 },
    });
  }
  if (/^\d+[.、]\s*/.test(text)) {
    return new Paragraph({
      text,
      spacing: { after: 120 },
    });
  }
  return new Paragraph({
    children: [new TextRun(text)],
    spacing: { after: 120 },
  });
}

async function createWordFromCodex(prompt, finalText) {
  const firstLine = finalText.split(/\r?\n/).map((line) => line.trim()).find(Boolean);
  const title = (firstLine || prompt).replace(/^#{1,3}\s+/, "").slice(0, 60);
  const children = [
    new Paragraph({
      text: title,
      heading: HeadingLevel.TITLE,
      spacing: { after: 240 },
    }),
    ...finalText
      .split(/\r?\n/)
      .filter((line, index) => index !== 0 || line.replace(/^#{1,3}\s+/, "").trim() !== title)
      .map(paragraphFromLine),
  ];

  const doc = new Document({
    creator: "Remote Codex Bridge",
    title,
    styles: {
      default: {
        document: {
          run: {
            font: "Microsoft YaHei",
            size: 22,
          },
          paragraph: {
            spacing: { line: 320 },
          },
        },
      },
    },
    sections: [{ children }],
  });

  fs.mkdirSync(path.join(ROOT, "outputs"), { recursive: true });
  const fileName = `${Date.now()}-${safeFilePart(title)}.docx`;
  const filePath = path.join(ROOT, "outputs", fileName);
  fs.writeFileSync(filePath, await Packer.toBuffer(doc));
  return filePath;
}

function runCodexBridge(prompt, options) {
  return new Promise((resolve) => {
    const stamp = Date.now();
    const outputFile = path.join(ROOT, "outputs", `codex-${stamp}.txt`);
    const logFile = path.join(ROOT, "outputs", `codex-${stamp}.log`);
    const codexBin = options.bin || "codex";
    const cwd = options.cwd || ROOT;
    const imageArgs = [];
    for (const imagePath of options.images || []) {
      if (imagePath && fs.existsSync(imagePath)) {
        imageArgs.push("-i", imagePath);
      }
    }
    const args = [
      "exec",
      ...imageArgs,
      "-C",
      cwd,
      "--skip-git-repo-check",
      "--sandbox",
      options.sandbox,
      "--output-last-message",
      outputFile,
      prompt,
    ];

    const child = spawn(codexBin, args, {
      cwd,
      windowsHide: true,
      shell: false,
      stdio: ["ignore", "pipe", "pipe"],
      env: {
        ...process.env,
        ...(options.home ? { CODEX_HOME: options.home } : {}),
      },
    });

    let stdout = "";
    let stderr = "";
    let settled = false;
    const timeout = setTimeout(() => {
      if (settled) return;
      settled = true;
      child.kill();
      resolve({
        ok: false,
        outputFile,
        logFile,
        error: `Codex Bridge timed out after ${options.timeoutMs}ms.`,
        stdout,
        stderr,
      });
    }, options.timeoutMs);

    child.stdout.on("data", (chunk) => {
      stdout += chunk.toString();
      fs.writeFileSync(logFile, stdout + stderr, "utf8");
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk.toString();
      fs.writeFileSync(logFile, stdout + stderr, "utf8");
    });
    child.on("error", (error) => {
      if (settled) return;
      settled = true;
      clearTimeout(timeout);
      resolve({ ok: false, outputFile, logFile, error: error.message, stdout, stderr });
    });
    child.on("close", (code) => {
      if (settled) return;
      settled = true;
      clearTimeout(timeout);
      let finalText = "";
      if (fs.existsSync(outputFile)) {
        finalText = fs.readFileSync(outputFile, "utf8").trim();
      }
      resolve({
        ok: code === 0 && Boolean(finalText),
        outputFile,
        logFile,
        finalText,
        error: code === 0 ? "" : `codex exec exited with code ${code}`,
        stdout,
        stderr,
      });
    });
  });
}

async function sendText(lark, client, chatId, text) {
  const body = {
    params: { receive_id_type: "chat_id" },
    data: {
      receive_id: chatId,
      msg_type: "text",
      content: JSON.stringify({ text }),
    },
  };

  if (client.im?.v1?.message?.create) {
    return client.im.v1.message.create(body);
  }
  return client.im.message.create(body);
}

async function sendFile(lark, client, chatId, filePath) {
  const fileName = path.basename(filePath);
  const ext = path.extname(filePath).toLowerCase();
  const fileTypeMap = {
    ".ppt": "ppt",
    ".pptx": "ppt",
    ".doc": "doc",
    ".docx": "doc",
    ".xls": "xls",
    ".xlsx": "xls",
    ".pdf": "pdf",
  };
  const upload = await client.im.file.create({
    data: {
      file_type: fileTypeMap[ext] || "stream",
      file_name: fileName,
      file: fs.readFileSync(filePath),
    },
  });

  const fileKey = upload?.data?.file_key || upload?.file_key;
  if (!fileKey) {
    throw new Error(`Feishu file upload did not return file_key: ${JSON.stringify(upload)}`);
  }

  const body = {
    params: { receive_id_type: "chat_id" },
    data: {
      receive_id: chatId,
      msg_type: "file",
      content: JSON.stringify({ file_key: fileKey }),
    },
  };
  if (client.im?.v1?.message?.create) {
    return client.im.v1.message.create(body);
  }
  return client.im.message.create(body);
}

async function downloadMessageFile(client, messageId, fileKey, fileName) {
  const uploadDir = path.join(ROOT, "uploads", String(Date.now()));
  fs.mkdirSync(uploadDir, { recursive: true });
  const safeName = safeUploadedName(fileName, "uploaded-file");
  const filePath = path.join(uploadDir, safeName);
  const resource = await client.im.v1.messageResource.get({
    path: {
      message_id: messageId,
      file_key: fileKey,
    },
    params: {
      type: "file",
    },
  });
  await resource.writeFile(filePath);
  return filePath;
}

async function downloadMessageImage(client, messageId, imageKey, content) {
  const uploadDir = path.join(ROOT, "uploads", "images", String(Date.now()));
  fs.mkdirSync(uploadDir, { recursive: true });
  const guessedExt = getImageExtension(content);
  const filePath = path.join(uploadDir, `photo${guessedExt}`);
  const resource = await client.im.v1.messageResource.get({
    path: {
      message_id: messageId,
      file_key: imageKey,
    },
    params: {
      type: "image",
    },
  });
  await resource.writeFile(filePath);
  const detectedExt = detectImageExtension(filePath, guessedExt);
  if (detectedExt !== guessedExt) {
    const detectedPath = path.join(uploadDir, `photo${detectedExt}`);
    fs.renameSync(filePath, detectedPath);
    return detectedPath;
  }
  return filePath;
}

async function handleUploadedFile({ lark, client, chatId, senderOpenId, messageId, content }) {
  const fileKey = content.file_key || content.fileKey;
  const fileName = content.file_name || content.fileName || content.name || "uploaded-file";
  const fileType = getOfficeFileTypeByName(fileName);
  if (!fileKey || !fileType) {
    await sendText(lark, client, chatId, "我收到了文件，但目前只支持 .doc/.docx/.ppt/.pptx。");
    return true;
  }

  const filePath = await downloadMessageFile(client, messageId, fileKey, fileName);
  latestOfficeFiles.set(getConversationKey(chatId, senderOpenId), {
    filePath,
    fileName: path.basename(filePath),
    fileType,
    receivedAt: Date.now(),
  });
  await sendText(lark, client, chatId, `已收到文件：${path.basename(filePath)}\n你可以继续发修改要求。`);
  return true;
}

async function handleUploadedImage({ lark, client, chatId, senderOpenId, messageId, content }) {
  const imageKey = content.image_key || content.imageKey;
  if (!imageKey) {
    await sendText(lark, client, chatId, "\u6211\u6536\u5230\u4e86\u56fe\u7247\uff0c\u4f46\u6ca1\u6709\u627e\u5230\u53ef\u4e0b\u8f7d\u7684 image_key\u3002");
    return true;
  }

  const filePath = await downloadMessageImage(client, messageId, imageKey, content);
  rememberLatestImage(chatId, senderOpenId, filePath);
  await sendText(
    lark,
    client,
    chatId,
    `\u5df2\u6536\u5230\u56fe\u7247\uff1a${path.basename(filePath)}\n\u4f60\u53ef\u4ee5\u7ee7\u7eed\u53d1\u5904\u7406\u8981\u6c42\uff0c\u6bd4\u5982\uff1a\u8bc6\u522b\u8fd9\u5f20\u56fe\u3001\u63d0\u53d6\u56fe\u7247\u6587\u5b57\u3001\u6839\u636e\u8fd9\u5f20\u56fe\u751f\u6210\u6587\u6863\u3002`
  );
  return true;
}

async function processCodexBridge({ lark, client, chatId, senderOpenId, messageId, codexBridge, prompt }) {
  try {
    const conversationKey = getConversationKey(chatId, senderOpenId);
    const latestFile =
      latestOfficeFiles.get(conversationKey) ||
      findLatestUploadedOfficeFile();
    const latestImage =
      latestImages.get(conversationKey) ||
      findLatestUploadedImage();
    const editMode = Boolean(latestFile && wantsEditExistingFile(prompt) && !/^\s*(新建|生成|创建)/.test(prompt));
    const fileType = editMode ? latestFile.fileType : getOfficeFileType(prompt);
    const fileMode = editMode || Boolean(fileType);
    const imageMode = Boolean(latestImage && wantsImageTask(prompt));
    const jobDir = fileMode ? createOfficeJobDir() : "";
    const startedAt = Date.now();
    let bridgePrompt = prompt;
    let inputName = "";
    if (editMode) {
      inputName = safeUploadedName(latestFile.fileName, `input.${latestFile.fileType}`);
      fs.copyFileSync(latestFile.filePath, path.join(jobDir, inputName));
      bridgePrompt = buildEditOfficeFilePrompt(prompt, inputName, latestFile.fileType);
    } else if (fileMode) {
      bridgePrompt = buildOfficeFilePrompt(prompt, fileType);
    }
    if (imageMode) {
      bridgePrompt = [
        bridgePrompt,
        "",
        "The user also uploaded an image. Use the attached image as input for this request.",
      ].join("\n");
    }
    const result = await runCodexBridge(bridgePrompt, {
      ...codexBridge,
      sandbox: fileMode ? "workspace-write" : codexBridge.sandbox,
      cwd: fileMode ? jobDir : ROOT,
      images: imageMode ? [latestImage.filePath] : [],
    });
    if (!result.ok) {
      await sendText(
        lark,
        client,
        chatId,
        `Codex Bridge 执行失败：\n${result.error}\n日志文件：${path.basename(result.logFile || "")}\n\n如果黑色窗口能联网但这里失败，可能需要从普通 PowerShell 启动 start-feishu.cmd。`
      );
      return;
    }
    if (fileMode) {
      const files = findNewOfficeFiles(jobDir, startedAt, fileType, inputName ? [inputName] : []);
      if (files.length === 0) {
        await sendText(
          lark,
          client,
          chatId,
          `Codex 已返回，但没有找到新生成的 .${fileType} 文件。\n${result.finalText}`
        );
        return;
      }
      await sendFile(lark, client, chatId, files[0]);
      rememberLatestOfficeFile(chatId, senderOpenId, files[0], fileType);
      return;
    }
    for (const chunk of splitText(result.finalText)) {
      await sendText(lark, client, chatId, chunk);
    }
  } catch (error) {
    await sendText(lark, client, chatId, `Codex Bridge 执行异常：\n${error.stack || error.message}`);
  } finally {
    if (messageId) activeCodexMessages.delete(messageId);
  }
}

async function handleMessage({ lark, client, allowedOpenIds, codexBridge }, data) {
  const payload = getEventPayload(data);
  if (isFromBot(payload)) return;

  const chatId = payload?.message?.chat_id;
  const messageId = payload?.message?.message_id || "";
  const senderOpenId = getSenderOpenId(payload);
  const msgType = payload?.message?.message_type || payload?.message?.msg_type || "";
  const content = parseMessageContent(payload?.message?.content);
  const text = getTextFromContent(payload?.message?.content).trim();

  console.log("[message]", {
    chatId,
    messageId,
    senderOpenId,
    text,
  });

  if (!chatId) return;
  if (allowedOpenIds.size > 0 && !allowedOpenIds.has(senderOpenId)) {
    await sendText(
      lark,
      client,
      chatId,
      `\u8fd9\u4e2a\u8d26\u53f7\u8fd8\u6ca1\u6709\u6388\u6743\u4f7f\u7528 Remote Codex\u3002\n\u4f60\u7684 open_id \u662f\uff1a${senderOpenId}`
    );
    return;
  }
  if (allowedOpenIds.size > 0 && !allowedOpenIds.has(senderOpenId)) {
    await sendText(
      lark,
      client,
      chatId,
      `这个账号还没有授权使用 Remote Codex。\n你的 open_id 是：${senderOpenId}`
    );
    return;
  }
  if (msgType === "image" || content.image_key || content.imageKey) {
    await handleUploadedImage({ lark, client, chatId, senderOpenId, messageId, content });
    return;
  }
  if (msgType === "file" || content.file_key || content.fileKey) {
    await handleUploadedFile({ lark, client, chatId, senderOpenId, messageId, content });
    return;
  }
  if (!text) {
    await sendText(lark, client, chatId, "我收到了消息，但暂时只处理文本指令。");
    return;
  }

  if (allowedOpenIds.size > 0 && !allowedOpenIds.has(senderOpenId)) {
    await sendText(
      lark,
      client,
      chatId,
      `这个账号还没有授权使用 Remote Codex。\n你的 open_id 是：${senderOpenId}`
    );
    return;
  }

  if (/^(测试|test|ping)$/i.test(text)) {
    await sendText(
      lark,
      client,
      chatId,
      `连接成功。我已经能收到你的飞书指令。\n你的 open_id 是：${senderOpenId || "unknown"}`
    );
    return;
  }

  if (codexBridge.enabled) {
    const prompt = stripCodexBridgePrefix(text);
    if (!prompt) {
      await sendText(lark, client, chatId, "请发送要处理的内容。");
      return;
    }
    if (messageId && activeCodexMessages.has(messageId)) {
      console.log("[codex] duplicate message ignored", messageId);
      return;
    }
    if (messageId) activeCodexMessages.add(messageId);
    processCodexBridge({ lark, client, chatId, senderOpenId, messageId, codexBridge, prompt });
    return;
  }

  await sendText(lark, client, chatId, "Codex Bridge 当前已关闭，请检查 CODEX_BRIDGE_ENABLED。");
}

async function main() {
  loadEnvFile(path.join(ROOT, ".env.local"));
  loadEnvFile(path.join(ROOT, ".env"));

  let lark;
  try {
    lark = require("@larksuiteoapi/node-sdk");
  } catch (error) {
    throw new Error(
      "Missing dependency @larksuiteoapi/node-sdk. Install it with npm install before starting the Feishu agent."
    );
  }

  const config = {
    appId: requireEnv("FEISHU_APP_ID"),
    appSecret: requireEnv("FEISHU_APP_SECRET"),
    appType: lark.AppType?.SelfBuild,
    domain: lark.Domain?.Feishu,
    loggerLevel: lark.LoggerLevel?.info,
  };

  const client = new lark.Client(config);
  const wsClient = new lark.WSClient({
    ...config,
    autoReconnect: true,
  });

  const allowedOpenIds = parseAllowedOpenIds(process.env.FEISHU_ALLOWED_OPEN_IDS);
  const codexBridge = {
    enabled: process.env.CODEX_BRIDGE_ENABLED !== "false",
    bin: process.env.CODEX_BRIDGE_BIN || "",
    home: process.env.CODEX_BRIDGE_HOME || "",
    sandbox: process.env.CODEX_BRIDGE_SANDBOX || "read-only",
    timeoutMs: Number(process.env.CODEX_BRIDGE_TIMEOUT_MS || 300000),
  };

  const dispatcherOptions = {};
  if (process.env.FEISHU_ENCRYPT_KEY) dispatcherOptions.encryptKey = process.env.FEISHU_ENCRYPT_KEY;
  if (process.env.FEISHU_VERIFICATION_TOKEN) {
    dispatcherOptions.verificationToken = process.env.FEISHU_VERIFICATION_TOKEN;
  }

  const eventDispatcher = new lark.EventDispatcher(dispatcherOptions).register({
    "im.message.receive_v1": (data) =>
      handleMessage({ lark, client, allowedOpenIds, codexBridge }, data),
  });

  console.log("Remote Codex Feishu Agent starting...");
  console.log(
    allowedOpenIds.size > 0
      ? `Allowed open_ids: ${Array.from(allowedOpenIds).join(", ")}`
      : "Allowed open_ids is empty. First run will accept messages and print your open_id."
  );
  console.log(
    codexBridge.enabled
      ? `Codex Bridge enabled. Bin: ${codexBridge.bin || "codex"}. Home: ${codexBridge.home || "default"}. Sandbox: ${codexBridge.sandbox}. Timeout: ${codexBridge.timeoutMs}ms.`
      : "Codex Bridge disabled."
  );
  console.log(`Cleanup retention: ${process.env.CLEANUP_RETENTION_DAYS || 7} days.`);
  scheduleCleanup(process.env.CLEANUP_RETENTION_DAYS || 7);
  wsClient.start({ eventDispatcher });
  console.log("Remote Codex Feishu Agent connected. Send '测试' to your bot.");
}

if (require.main === module) {
  main().catch((error) => {
    console.error(error);
    process.exit(1);
  });
}
