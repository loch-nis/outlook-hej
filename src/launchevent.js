/*
 * Hej-hilsen: puts "Hej <fornavn>" at the top of a mail being composed and
 * keeps it in step with the To field until the user edits it.
 *
 * Classic Outlook on Windows loads this file directly in a JavaScript-only
 * runtime, so it must stay a single file with no imports and no DOM access.
 */

const CONFIG = {
  salutation: "Hej",
  punctuation: ",", // "" gives "Hej Anne"
  and: "og",
  everyone: "alle",
  maxNames: 3, // more To recipients than this gives "Hej alle"
  // On reply, classic Outlook fires both events at once in separate runtimes;
  // the recipients handler waits so the compose handler inserts first.
  replyDebounceMs: 1000,
  callTimeoutMs: 10000,
  debug: true, // shows what each event did in a notice bar
};

// A first body line starting with one of these is the user's own greeting.
const GREETING_WORDS = [
  "hej", "hejsa", "hey", "hi", "hello", "hallo", "halløj", "kære", "dear",
  "goddag", "godmorgen", "dav", "davs", "morn", "hola",
];
const HONORIFICS = new Set(["dr", "mr", "mrs", "ms", "miss", "prof", "hr", "fru", "frk", "sir"]);
const ROLE_MAILBOXES = new Set([
  "info", "kontakt", "contact", "support", "hello", "hej", "mail", "post", "salg", "sales",
  "admin", "faktura", "invoice", "invoices", "billing", "noreply", "no-reply", "donotreply",
  "office", "kontor", "team", "job", "jobs", "service", "kundeservice", "booking", "ordre",
  "order", "orders", "regnskab", "bogholderi", "accounting", "marketing", "webmaster", "postmaster",
]);
const SESSION_KEY = "hejGreeting";

// ---------- names ----------

function titleCase(word) {
  return word
    .split("-")
    .map((part) => part.charAt(0).toLocaleUpperCase("da") + part.slice(1).toLocaleLowerCase("da"))
    .join("-");
}

// Keeps mixed-case names such as "McKenzie" untouched; fixes "ANNE" and "anne".
function normalizeCase(word) {
  const isUpper = word === word.toLocaleUpperCase("da");
  const isLower = word === word.toLocaleLowerCase("da");
  return isUpper || isLower ? titleCase(word) : word;
}

function isNameToken(token) {
  return /^\p{L}[\p{L}'’-]*\p{L}$/u.test(token);
}

function nameFromDisplay(display) {
  let name = display
    .replace(/^["'\s]+|["'\s]+$/g, "")
    .replace(/\([^)]*\)|\[[^\]]*\]|<[^>]*>/g, " ")
    .split(/\s[|–—-]\s|\s\/\s/)[0]
    .trim();
  if (name.includes(",")) {
    // "Hansen, Anne" puts the first name after the comma; "Anne Hansen, CEO" before it.
    const [before, after] = name.split(",", 2).map((s) => s.trim());
    name = before.split(/\s+/).length === 1 && after ? after : before;
  }
  const tokens = name.split(/\s+/).filter((t) => !HONORIFICS.has(t.toLowerCase().replace(/\.$/, "")));
  const first = tokens[0] || "";
  return isNameToken(first) ? normalizeCase(first) : "";
}

// Only "anne.hansen@" style addresses reveal a first name; "anne@" or "ah@" are guesses.
function nameFromAddress(address) {
  const local = (address || "").split("@")[0].split("+")[0].toLowerCase();
  if (ROLE_MAILBOXES.has(local)) return "";
  const parts = local.split(/[._]/);
  if (parts.length < 2 || !isNameToken(parts[0])) return "";
  return titleCase(parts[0]);
}

function firstNameOf(recipient) {
  const display = (recipient.displayName || "").trim();
  if (display && !display.includes("@")) return nameFromDisplay(display);
  return nameFromAddress(recipient.emailAddress || display);
}

function joinNames(names) {
  if (names.length === 1) return names[0];
  return names.slice(0, -1).join(", ") + " " + CONFIG.and + " " + names[names.length - 1];
}

function buildGreeting(recipients, selfAddress) {
  const self = (selfAddress || "").toLowerCase();
  const others = (recipients || []).filter((r) => (r.emailAddress || "").toLowerCase() !== self);
  if (others.length === 0) return "";

  const end = CONFIG.punctuation;
  const everyone = CONFIG.salutation + " " + CONFIG.everyone + end;
  const hasGroup = others.some((r) => r.recipientType === "distributionList");
  if (hasGroup || others.length > CONFIG.maxNames) return everyone;

  const names = others.map(firstNameOf);
  if (names.length === 1) return CONFIG.salutation + (names[0] ? " " + names[0] : "") + end;
  if (names.some((n) => !n)) return everyone;
  return CONFIG.salutation + " " + joinNames(names) + end;
}

// ---------- body text ----------

const NAMED_ENTITIES = {
  nbsp: " ", amp: "&", lt: "<", gt: ">", quot: '"', apos: "'",
  aelig: "æ", AElig: "Æ", oslash: "ø", Oslash: "Ø", aring: "å", Aring: "Å",
  auml: "ä", Auml: "Ä", ouml: "ö", Ouml: "Ö", uuml: "ü", Uuml: "Ü",
  eacute: "é", Eacute: "É", egrave: "è", aacute: "á", oacute: "ó", iacute: "í", ntilde: "ñ", szlig: "ß",
};

function decodeEntities(s) {
  return s.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (m, code) => {
    if (code[0] === "#") {
      const n = code[1] === "x" || code[1] === "X" ? parseInt(code.slice(2), 16) : parseInt(code.slice(1), 10);
      return n > 0 && n <= 0x10ffff ? String.fromCodePoint(n) : m;
    }
    return NAMED_ENTITIES[code] !== undefined ? NAMED_ENTITIES[code] : m;
  });
}

function bodyStart(html) {
  const m = /<body\b[^>]*>/i.exec(html);
  return m ? m.index + m[0].length : 0;
}

function htmlToText(html) {
  return decodeEntities(
    html
      .slice(bodyStart(html))
      .replace(/<!--[\s\S]*?-->/g, "")
      .replace(/<(style|script|head)\b[\s\S]*?<\/\1>/gi, "")
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<\/(p|div|li|tr|h[1-6]|table|blockquote)>/gi, "\n")
      .replace(/<[^>]+>/g, "")
  );
}

function firstLine(text) {
  const line = text.split(/\r?\n/).find((l) => l.replace(/[\s\u00a0]+/g, "") !== "");
  return line ? line.replace(/[\s\u00a0]+/g, " ").trim() : "";
}

function startsWithGreeting(line) {
  const lower = line.toLocaleLowerCase("da");
  const word = lower.split(/[\s,!.:;]+/)[0];
  return GREETING_WORDS.includes(word) || lower.startsWith("god morgen") || lower.startsWith("go' morgen");
}

function escapeHtml(s) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// Matches the greeting as it may come back from the client: spaces as &nbsp;,
// letters such as ø as &oslash; or &#248;.
function greetingPattern(greeting) {
  const parts = Array.from(greeting).map((ch) => {
    if (/\s/.test(ch)) return "(?:\\s|&nbsp;|&#160;|&#xa0;)+";
    const code = ch.codePointAt(0);
    const alternatives = [escapeRegExp(ch), "&#" + code + ";", "&#x" + code.toString(16) + ";"];
    const named = Object.keys(NAMED_ENTITIES).find((k) => NAMED_ENTITIES[k] === ch);
    if (named) alternatives.push("&" + named + ";");
    return "(?:" + alternatives.join("|") + ")";
  });
  return new RegExp(parts.join(""), "gi");
}

// Replaces the greeting only where it is the first visible text of the body.
function replaceGreetingHtml(html, previous, next) {
  const start = bodyStart(html);
  const pattern = greetingPattern(previous);
  pattern.lastIndex = start;
  const match = pattern.exec(html);
  if (!match || firstLine(htmlToText(html.slice(start, match.index))) !== "") return null;
  return html.slice(0, match.index) + escapeHtml(next) + html.slice(match.index + match[0].length);
}

function replaceGreetingText(text, previous, next) {
  const index = text.indexOf(previous);
  if (index < 0 || firstLine(text.slice(0, index)) !== "") return null;
  return text.slice(0, index) + next + text.slice(index + previous.length);
}

function attribute(tag, name) {
  const m = new RegExp("\\s" + name + "\\s*=\\s*(\"[^\"]*\"|'[^']*'|[^\\s>]+)", "i").exec(tag);
  return m ? m[1].replace(/^["']|["']$/g, "") : "";
}

const EDITOR_CLASSES = ["MsoNormal", "elementToProof"];

function editorClasses(tag) {
  return attribute(tag, "class").split(/\s+/).filter((c) => EDITOR_CLASSES.includes(c));
}

function fontStyle(tag) {
  return attribute(tag, "style")
    .split(";")
    .filter((d) => /^\s*(font-family|font-size|color)\s*:/i.test(d))
    .join(";")
    .trim();
}

// The first paragraph the user would type into: a styled block before any visible text,
// skipping wrappers such as Word's <div class=WordSection1>.
function firstParagraphTag(html) {
  const start = bodyStart(html);
  const blocks = /<(p|div)\b[^>]*>/gi;
  blocks.lastIndex = start;
  let m;
  while ((m = blocks.exec(html)) && firstLine(htmlToText(html.slice(start, m.index))) === "") {
    if (editorClasses(m[0]).length || fontStyle(m[0])) return m[0];
  }
  return "";
}

// Copies the font of the body's first paragraph so the greeting matches what the user types.
function greetingHtml(html, greeting) {
  const tag = firstParagraphTag(html);
  const classes = editorClasses(tag);
  const style = fontStyle(tag);
  const isWord = classes.includes("MsoNormal");
  const name = isWord ? "p" : "div";
  const attrs =
    (classes.length ? ' class="' + classes.join(" ") + '"' : "") +
    (style ? ' style="' + style.replace(/"/g, "'") + '"' : "");
  const blank = isWord ? "&nbsp;" : "<br>";
  return "<" + name + attrs + ">" + escapeHtml(greeting) + "</" + name + ">" +
    "<" + name + attrs + ">" + blank + "</" + name + ">";
}

// ---------- Outlook ----------

// Rejects instead of hanging when Outlook never calls back, so the event still completes.
function officeCall(target, method, ...args) {
  return new Promise((resolve, reject) => {
    const timer = typeof setTimeout === "function"
      ? setTimeout(() => reject(new Error(method + ": no answer from Outlook")), CONFIG.callTimeoutMs)
      : null;
    target[method](...args, (result) => {
      if (timer) clearTimeout(timer);
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else reject(new Error(method + ": " + (result.error && (result.error.message || result.error.code))));
    });
  });
}

// Shows text in the compose window; event runtimes have no console the user can see.
function showNotice(key, text) {
  const item = Office.context.mailbox.item;
  if (!item || !item.notificationMessages) return;
  item.notificationMessages.replaceAsync(key, {
    type: "informationalMessage",
    message: ("Hej-hilsen: " + text).slice(0, 150),
    icon: "Icon.16x16",
    persistent: false,
  });
}

function showProblem(error) {
  showNotice("hejHilsenProblem", error && error.message ? error.message : String(error));
}

function trace(text) {
  if (CONFIG.debug) showNotice("hejHilsenDebug", text);
}

function delay(ms) {
  return new Promise((resolve) => (typeof setTimeout === "function" ? setTimeout(resolve, ms) : resolve()));
}

async function readPrevious(item) {
  if (!item.sessionData) return "";
  try {
    return (await officeCall(item.sessionData, "getAsync", SESSION_KEY)) || "";
  } catch (e) {
    return "";
  }
}

async function rememberGreeting(item, greeting) {
  if (item.sessionData) await officeCall(item.sessionData, "setAsync", SESSION_KEY, greeting);
}

// Where the user's own text ends: Outlook's signature or the quoted mail
// (web and Mac ids, classic Windows bookmark and reply header).
const SIGNATURE_OR_QUOTE =
  /id\s*=\s*["']?(?:x_)*(?:Signature|appendonsend|divRplyFwdMsg|mail-editor-reference-message-container)\b|name\s*=\s*["']?_MailAutoSig|border-top:\s*solid\s+#E1E1E1|<hr\b/i;

function typedText(html) {
  const rest = html.slice(bodyStart(html));
  const marker = rest.search(SIGNATURE_OR_QUOTE);
  const own = marker < 0 ? rest : rest.slice(0, rest.lastIndexOf("<", marker));
  return firstLine(htmlToText(own));
}

function hasInlineImages(html) {
  return /<img\b[^>]*(?:src\s*=\s*["']?(?:cid:|data:)|AttachmentByCid)/i.test(html);
}

// In Outlook on the web and new Outlook, limits body reads and writes on a reply to
// the part being written instead of the whole thread. Other clients ignore it.
function bodyModeOption() {
  const modes = Office.MailboxEnums && Office.MailboxEnums.BodyMode;
  return modes ? { bodyMode: modes.HostConfig } : {};
}

async function composeType(item) {
  if (!item.getComposeTypeAsync) return "newMail";
  try {
    return (await officeCall(item, "getComposeTypeAsync")).composeType;
  } catch (e) {
    return "newMail";
  }
}

async function applyGreeting(fromRecipientsEvent) {
  const mailbox = Office.context.mailbox;
  const item = mailbox.item;
  const type = await composeType(item);
  if (fromRecipientsEvent && type === "reply") await delay(CONFIG.replyDebounceMs);
  const to = await officeCall(item.to, "getAsync");
  const greeting = buildGreeting(to, mailbox.userProfile && mailbox.userProfile.emailAddress);
  if (!greeting) return "no To recipients";

  const isHtml = (await officeCall(item.body, "getTypeAsync")) === Office.CoercionType.Html;
  const coercionType = isHtml ? Office.CoercionType.Html : Office.CoercionType.Text;
  const body = await officeCall(item.body, "getAsync", coercionType, bodyModeOption());
  const line = firstLine(isHtml ? htmlToText(body) : body);
  const previous = await readPrevious(item);

  if (previous && line === previous) {
    if (previous === greeting) return "unchanged";
    // Rewriting a reply or forward can lose the quoted mail's inline images (office-js#6808, #6944).
    if (type !== "newMail" && hasInlineImages(body)) return "kept (quoted inline images)";
    const updated = isHtml
      ? replaceGreetingHtml(body, previous, greeting)
      : replaceGreetingText(body, previous, greeting);
    if (updated === null) return "greeting not found";
    await officeCall(item.body, "setAsync", updated, Object.assign({ coercionType }, bodyModeOption()));
    await rememberGreeting(item, greeting);
    return "updated to " + greeting;
  }
  // The user wrote or edited a greeting; leave the body alone from here on.
  if (startsWithGreeting(line)) return "left the user's greeting alone";

  const content = isHtml ? greetingHtml(body, greeting) : greeting + "\n\n";
  // Before the user has typed anything, inserting at the selection (the top of the body)
  // leaves the cursor below the greeting; prependAsync leaves it above.
  let method = "";
  if (isHtml && !typedText(body)) {
    try {
      await officeCall(item.body, "setSelectedDataAsync", content, { coercionType });
      method = "setSelectedDataAsync";
    } catch (error) {
      console.error("Hej-hilsen: " + error.message);
    }
  }
  if (!method) {
    await officeCall(item.body, "prependAsync", content, { coercionType });
    method = "prependAsync";
  }
  await rememberGreeting(item, greeting);
  return "inserted " + greeting + " (" + type + ", " + method + ")";
}

// Events can overlap (reply opening, pasting several addresses); run them one at a time.
let queue = Promise.resolve();

function handleEvent(event, fromRecipientsEvent) {
  const name = fromRecipientsEvent ? "recipients changed" : "compose";
  trace(name + ": started");
  queue = queue
    .then(() => applyGreeting(fromRecipientsEvent))
    .then((outcome) => trace(name + ": " + outcome))
    .catch((error) => {
      console.error("Hej-hilsen: " + (error && error.message));
      showProblem(error);
    })
    .then(() => event.completed());
}

function onNewMessageComposeHandler(event) {
  handleEvent(event, false);
}

function onMessageRecipientsChangedHandler(event) {
  if (event.changedRecipientFields && !event.changedRecipientFields.to) {
    event.completed();
    return;
  }
  handleEvent(event, true);
}

if (typeof Office !== "undefined" && Office.actions) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);
}

if (typeof module !== "undefined") {
  module.exports = {
    CONFIG,
    firstNameOf,
    buildGreeting,
    htmlToText,
    firstLine,
    startsWithGreeting,
    replaceGreetingHtml,
    replaceGreetingText,
    greetingHtml,
    typedText,
    onNewMessageComposeHandler,
    onMessageRecipientsChangedHandler,
  };
}
