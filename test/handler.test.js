// Drives the event handlers against a fake Outlook compose item.
const test = require("node:test");
const assert = require("node:assert/strict");
const g = require("../src/launchevent.js");

g.CONFIG.recipientsDebounceMs = 30;

// A second copy of the script, like classic Outlook starting a fresh runtime per event.
function freshRuntime() {
  const file = require.resolve("../src/launchevent.js");
  delete require.cache[file];
  const copy = require(file);
  copy.CONFIG.recipientsDebounceMs = 30;
  return copy;
}

const person = (displayName, emailAddress) => ({ displayName, emailAddress, recipientType: "externalUser" });
const ANNE = person("Anne Hansen", "anne@example.com");
const PETER = person("Peter Holm", "peter@example.com");
const SOEREN = person("Søren Ørsted", "soeren@example.com");

const OWA_NEW =
  '<html><head><style>p{margin-top:0;margin-bottom:0}</style></head><body dir="ltr">' +
  '<div class="elementToProof" style="font-family: Aptos, Aptos_EmbeddedFont, Calibri, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);"><br></div>' +
  '<div id="Signature"><div>Med venlig hilsen</div><div>Nis</div></div></body></html>';
const OWA_REPLY =
  '<html><head></head><body dir="ltr">' +
  '<div class="elementToProof" style="font-family: Aptos; font-size: 12pt; color: rgb(0, 0, 0);"><br></div>' +
  '<div id="appendonsend"></div><hr><div id="divRplyFwdMsg"><b>Fra:</b> Anne Hansen<br><b>Emne:</b> Møde</div>' +
  "<div>Hej Nis</div><div>Kan vi mødes?</div></body></html>";

// `encode` mimics clients that store ø as &oslash; after an insert.
function fakeOutlook({ to = [], body = OWA_NEW, type = "html", encode = false, sessionData = true, bodyMode = false } = {}) {
  const state = { to, body, type, session: {}, writes: 0, options: [] };
  const ok = (cb, value) => cb({ status: "succeeded", value });
  const last = (args) => args[args.length - 1];
  const stored = (html) => (encode ? html.replace(/ø/g, "&oslash;").replace(/Ø/g, "&Oslash;") : html);
  const item = {
    to: { getAsync: (cb) => ok(cb, state.to) },
    body: {
      getTypeAsync: (cb) => ok(cb, state.type),
      getAsync: (...args) => {
        state.options.push(args[1]);
        ok(last(args), state.body);
      },
      prependAsync: (data, options, cb) => {
        state.writes++;
        const at = state.type === "html" ? state.body.search(/<body[^>]*>/i) : -1;
        if (at < 0) state.body = stored(data) + state.body;
        else {
          const end = state.body.indexOf(">", at) + 1;
          state.body = state.body.slice(0, end) + stored(data) + state.body.slice(end);
        }
        ok(cb);
      },
      setAsync: (data, options, cb) => {
        state.writes++;
        state.options.push(options);
        state.body = stored(data);
        ok(cb);
      },
    },
    sessionData: sessionData && {
      getAsync: (key, cb) =>
        key in state.session ? ok(cb, state.session[key]) : cb({ status: "failed", error: { code: 9057 } }),
      setAsync: (key, value, cb) => {
        state.session[key] = value;
        ok(cb);
      },
    },
  };
  global.Office = {
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    CoercionType: { Html: "html", Text: "text" },
    MailboxEnums: bodyMode ? { BodyMode: { FullBody: 0, HostConfig: 1 } } : {},
    context: { mailbox: { item, userProfile: { emailAddress: "me@example.com" } } },
  };
  return state;
}

function fire(handler, changedRecipientFields) {
  return new Promise((resolve) => handler({ changedRecipientFields, completed: resolve }));
}
const compose = (runtime = g) => fire(runtime.onNewMessageComposeHandler);
const toChanged = (runtime = g) => fire(runtime.onMessageRecipientsChangedHandler, { to: true, cc: false, bcc: false });
const top = (state) => g.firstLine(state.type === "html" ? g.htmlToText(state.body) : state.body);

test("new mail: greeting follows the To field", async () => {
  const state = fakeOutlook();
  await compose();
  assert.equal(state.writes, 0, "nothing to greet before recipients are added");

  state.to = [ANNE];
  await toChanged();
  assert.equal(top(state), "Hej Anne");
  assert.match(state.body, /<div class="elementToProof" style="font-family: Aptos[^"]*">Hej Anne<\/div>/);

  state.to = [ANNE, PETER];
  await toChanged();
  assert.equal(top(state), "Hej Anne og Peter");

  state.to = [PETER];
  await toChanged();
  assert.equal(top(state), "Hej Peter");

  state.to = [PETER, ANNE, SOEREN, person("Mette Lund", "m@example.com")];
  await toChanged();
  assert.equal(top(state), "Hej alle");
  assert.equal((state.body.match(/Hej/g) || []).length, 1, "only one greeting");
  assert.match(state.body, /Med venlig hilsen/);
});

test("emptying the To field keeps the last greeting", async () => {
  const state = fakeOutlook({ to: [ANNE] });
  await toChanged();
  state.to = [];
  await toChanged();
  assert.equal(top(state), "Hej Anne");
});

test("reply: greeting goes above the quoted mail, once, even with overlapping events", async () => {
  const state = fakeOutlook({ to: [ANNE], body: OWA_REPLY });
  await Promise.all([compose(), toChanged()]);
  assert.equal(top(state), "Hej Anne");
  assert.equal((state.body.match(/Hej Anne/g) || []).length, 1);
  assert.match(state.body, /<div>Hej Nis<\/div><div>Kan vi mødes\?<\/div>/, "quoted mail untouched");
});

test("a greeting the user edited is left alone", async () => {
  const state = fakeOutlook({ to: [ANNE] });
  await toChanged();
  state.body = state.body.replace("Hej Anne", "Hej Anne!");
  state.to = [ANNE, PETER];
  await toChanged();
  assert.equal(top(state), "Hej Anne!");
});

test("a greeting the user wrote before adding recipients is left alone", async () => {
  const state = fakeOutlook({ body: "<body><div>Kære Anne</div><div>Tak for sidst</div></body>" });
  state.to = [ANNE];
  await toChanged();
  assert.equal(state.writes, 0);
});

test("text typed before adding recipients gets the greeting above it", async () => {
  const state = fakeOutlook({ body: "<body><div>Tak for sidst</div></body>" });
  state.to = [ANNE];
  await toChanged();
  assert.equal(top(state), "Hej Anne");
  assert.match(g.htmlToText(state.body), /Hej Anne\n+Tak for sidst/);
});

test("names stored as entities are still updated", async () => {
  const state = fakeOutlook({ to: [SOEREN], encode: true });
  await toChanged();
  assert.match(state.body, /Hej S&oslash;ren/);
  state.to = [SOEREN, ANNE];
  await toChanged();
  assert.equal(top(state), "Hej Søren og Anne");
});

test("plain-text body", async () => {
  const state = fakeOutlook({ to: [ANNE], type: "text", body: "\n\nMvh Nis" });
  await toChanged();
  assert.equal(state.body, "Hej Anne\n\n\n\nMvh Nis");
  state.to = [ANNE, PETER];
  await toChanged();
  assert.equal(state.body, "Hej Anne og Peter\n\n\n\nMvh Nis");
});

test("Cc and Bcc changes are ignored", async () => {
  const state = fakeOutlook({ to: [ANNE] });
  await fire(g.onMessageRecipientsChangedHandler, { to: false, cc: true, bcc: false });
  assert.equal(state.writes, 0);
});

test("without sessionData the greeting is inserted once and never duplicated", async () => {
  const state = fakeOutlook({ to: [ANNE], sessionData: false });
  await toChanged();
  state.to = [ANNE, PETER];
  await toChanged();
  assert.equal(top(state), "Hej Anne");
  assert.equal(state.writes, 1);
});

test("an Outlook error still completes the event", async () => {
  const state = fakeOutlook({ to: [ANNE] });
  Office.context.mailbox.item.body.getTypeAsync = (cb) => cb({ status: "failed", error: { code: 5001 } });
  const original = console.error;
  console.error = () => {};
  try {
    await toChanged();
  } finally {
    console.error = original;
  }
  assert.equal(state.writes, 0);
});

test("reply in classic Outlook: separate runtimes for both events still give one greeting", async () => {
  const state = fakeOutlook({ to: [ANNE], body: OWA_REPLY });
  await Promise.all([compose(freshRuntime()), toChanged(freshRuntime())]);
  assert.equal((state.body.match(/Hej Anne/g) || []).length, 1);
});

test("a body with attached inline images is never rewritten", async () => {
  const withLogo = OWA_NEW.replace("<div>Nis</div>", '<div>Nis</div><img src="cid:logo.png" data-imagetype="AttachmentByCid">');
  const state = fakeOutlook({ to: [ANNE], body: withLogo });
  await toChanged();
  assert.equal(top(state), "Hej Anne", "first insert uses prependAsync, which is safe");
  state.to = [ANNE, PETER];
  await toChanged();
  assert.equal(top(state), "Hej Anne");
  assert.match(state.body, /cid:logo.png/);
});

test("a linked logo does not block updates", async () => {
  const withLogo = OWA_NEW.replace("<div>Nis</div>", '<div>Nis</div><img src="https://example.com/logo.png">');
  const state = fakeOutlook({ to: [ANNE], body: withLogo });
  await toChanged();
  state.to = [ANNE, PETER];
  await toChanged();
  assert.equal(top(state), "Hej Anne og Peter");
});

test("reads and writes only the current reply where the client supports it", async () => {
  const state = fakeOutlook({ to: [ANNE], bodyMode: true });
  await toChanged();
  state.to = [ANNE, PETER];
  await toChanged();
  assert.ok(state.options.length >= 3);
  for (const options of state.options) assert.equal(options.bodyMode, 1);
});
