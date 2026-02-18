const inputUrl = document.getElementById("inputUrl");
const outputUrl = document.getElementById("outputUrl");
const decodeBtn = document.getElementById("decodeBtn");
const status = document.getElementById("status");
const appLink = document.getElementById("appLink");

const OFFICE_PROTOCOLS = {
  doc: { prefix: "ms-word:ofv|u|", app: "Word" },
  docx: { prefix: "ms-word:ofv|u|", app: "Word" },
  docm: { prefix: "ms-word:ofv|u|", app: "Word" },
  xls: { prefix: "ms-excel:ofv|u|", app: "Excel" },
  xlsx: { prefix: "ms-excel:ofv|u|", app: "Excel" },
  xlsm: { prefix: "ms-excel:ofv|u|", app: "Excel" },
  ppt: { prefix: "ms-powerpoint:ofv|u|", app: "PowerPoint" },
  pptx: { prefix: "ms-powerpoint:ofv|u|", app: "PowerPoint" },
  pptm: { prefix: "ms-powerpoint:ofv|u|", app: "PowerPoint" }
};

/**
 * SharePoint特有の ":w:/s" などの区切りを削除します。
 * @param {string} path - URLのパス部分。
 * @returns {string} 置換後のパス。
 */
const stripSharePointMarker = (path) => {
  return path.replace(/\/:\w:\/\w\//i, "/");
};

/**
 * SharePoint URLを正規化してデコードします。
 * @param {string} rawUrl - 入力されたURL文字列。
 * @returns {string} デコード済みURL。
 */
const normalizeSharePointUrl = (rawUrl) => {
  const trimmed = rawUrl.trim();
  if (!trimmed) {
    throw new Error("URLが空です。入力してください。");
  }

  let url;
  try {
    url = new URL(trimmed);
  } catch {
    throw new Error("URLの形式が正しくありません。");
  }

  const hostname = url.hostname.toLowerCase();
  if (!hostname.endsWith(".sharepoint.com")) {
    throw new Error("sharepoint.comのURLのみ対応しています。");
  }

  const cleanedPath = stripSharePointMarker(url.pathname);
  const decodedPath = decodeURIComponent(cleanedPath);
  return `${url.protocol}//${url.host}${decodedPath}`;
};

/**
 * Officeアプリ用の起動リンクを組み立てます。
 * @param {string} normalizedUrl - 正規化済みURL。
 * @returns {{href: string, app: string} | null} 生成結果。対象拡張子が無い場合はnull。
 */
const buildOfficeLink = (normalizedUrl) => {
  const url = new URL(normalizedUrl);
  const match = url.pathname.toLowerCase().match(/\.([a-z0-9]+)$/);
  const extension = match ? match[1] : "";
  const protocol = OFFICE_PROTOCOLS[extension];
  if (!protocol) {
    return null;
  }
  return {
    href: `${protocol.prefix}${normalizedUrl}`,
    app: protocol.app
  };
};

/**
 * 画面上のステータスメッセージを更新します。
 * @param {string} message - 表示する文言。
 * @param {boolean} [isError=false] - エラー表示にするかどうか。
 */
const setStatus = (message, isError = false) => {
  status.textContent = message;
  status.style.color = isError ? "#b42318" : "#2b7a0b";
};

/**
 * 入力URLをデコードし、クリップボードへコピーします。
 * @returns {Promise<void>} 処理完了を待機するPromise。
 */
const handleDecode = async () => {
  setStatus("");
  try {
    const normalizedUrl = normalizeSharePointUrl(inputUrl.value);
    const officeLink = buildOfficeLink(normalizedUrl);
    outputUrl.value = normalizedUrl;
    outputUrl.focus();
    outputUrl.select();
    await navigator.clipboard.writeText(normalizedUrl);
    if (officeLink) {
      appLink.href = officeLink.href;
      appLink.textContent = `アプリで読み取り専用で開く（${officeLink.app}）`;
      appLink.hidden = false;
    } else {
      appLink.hidden = true;
      appLink.removeAttribute("href");
    }
    setStatus("デコードしてクリップボードにコピーしました。");
  } catch (error) {
    outputUrl.value = "";
    appLink.hidden = true;
    appLink.removeAttribute("href");
    setStatus(error.message, true);
  }
};

decodeBtn.addEventListener("click", handleDecode);
