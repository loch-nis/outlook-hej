const test = require("node:test");
const assert = require("node:assert/strict");
const g = require("../src/launchevent.js");

const r = (displayName, emailAddress = "x@example.com", recipientType = "externalUser") => ({
  displayName,
  emailAddress,
  recipientType,
});

test("first name from display names", () => {
  const cases = [
    ["Anne Hansen", "Anne"],
    ["Hansen, Anne", "Anne"],
    ["Hansen, Anne Marie", "Anne"],
    ["Anne Hansen, CEO", "Anne"],
    ["ANNE HANSEN", "Anne"],
    ["anne hansen", "Anne"],
    ["Anne-Marie Jensen", "Anne-Marie"],
    ["ANNE-MARIE JENSEN", "Anne-Marie"],
    ["Søren Ørsted", "Søren"],
    ["ÅSE ÆRØ", "Åse"],
    ["Dr. Anne Hansen", "Anne"],
    ["Anne Hansen (Acme A/S)", "Anne"],
    ["[EXT] Anne Hansen", "Anne"],
    ["Anne Hansen | Acme", "Anne"],
    ["'Anne Hansen'", "Anne"],
    ["Mary O'Neil", "Mary"],
    ["McKenzie Smith", "McKenzie"],
    ["A. Hansen", ""],
    ["Acme A/S", "Acme"],
  ];
  for (const [display, expected] of cases) {
    assert.equal(g.firstNameOf(r(display)), expected, display);
  }
});

test("first name from bare addresses", () => {
  assert.equal(g.firstNameOf(r("anne.hansen@firma.dk", "anne.hansen@firma.dk")), "Anne");
  assert.equal(g.firstNameOf(r("", "soeren_madsen@firma.dk")), "Soeren");
  assert.equal(g.firstNameOf(r("", "anne@firma.dk")), "");
  assert.equal(g.firstNameOf(r("info@firma.dk", "info@firma.dk")), "");
  assert.equal(g.firstNameOf(r("", "a.hansen@firma.dk")), "");
  assert.equal(g.firstNameOf(r("", "no-reply@firma.dk")), "");
});

test("greeting for one, two, three and many recipients", () => {
  assert.equal(g.buildGreeting([r("Anne Hansen")]), "Hej Anne,");
  assert.equal(g.buildGreeting([r("Anne Hansen"), r("Peter Holm")]), "Hej Anne og Peter,");
  assert.equal(g.buildGreeting([r("Anne Hansen"), r("Peter Holm"), r("Mette Lund")]), "Hej Anne, Peter og Mette,");
  assert.equal(g.buildGreeting([r("A B"), r("Cc Dd"), r("Ee Ff"), r("Gg Hh")]), "Hej alle,");
  assert.equal(g.buildGreeting([]), "");
});

test("greeting without a usable name", () => {
  assert.equal(g.buildGreeting([r("info@firma.dk", "info@firma.dk")]), "Hej,");
  assert.equal(g.buildGreeting([r("Anne Hansen"), r("info@firma.dk", "info@firma.dk")]), "Hej alle,");
  assert.equal(g.buildGreeting([r("Salg", "salg@firma.dk", "distributionList")]), "Hej alle,");
});

test("own address is left out of the greeting", () => {
  const me = r("Me Myself", "me@example.com");
  assert.equal(g.buildGreeting([me], "ME@example.com"), "");
  assert.equal(g.buildGreeting([me, r("Anne Hansen")], "me@example.com"), "Hej Anne,");
});

test("punctuation setting", () => {
  g.CONFIG.punctuation = "";
  try {
    assert.equal(g.buildGreeting([r("Anne Hansen")]), "Hej Anne");
  } finally {
    g.CONFIG.punctuation = ",";
  }
});

test("user greetings are recognised", () => {
  for (const line of ["Hej Anne", "hej", "Kære Anne,", "Hi Anne", "Godmorgen alle", "God morgen", "Halløj", "Dear Sir"]) {
    assert.ok(g.startsWithGreeting(line), line);
  }
  for (const line of ["Med venlig hilsen", "Hejre er fugle", "Fra: Anne", "Tak for sidst", ""]) {
    assert.ok(!g.startsWithGreeting(line), line);
  }
});

test("first visible line of an HTML body", () => {
  const html =
    '<html><head><style>p{margin:0}</style></head><body><div class="elementToProof"><br></div>' +
    "<div>S&oslash;ren&nbsp;og&#160;Anne</div></body></html>";
  assert.equal(g.firstLine(g.htmlToText(html)), "Søren og Anne");
});

test("replaces an entity-encoded greeting only at the top", () => {
  const html = "<body><div>Hej S&oslash;ren</div><div><br></div><div>Hej Søren igen</div></body>";
  assert.equal(
    g.replaceGreetingHtml(html, "Hej Søren", "Hej Søren og Anne"),
    "<body><div>Hej Søren og Anne</div><div><br></div><div>Hej Søren igen</div></body>"
  );
  assert.equal(g.replaceGreetingHtml("<body><div>Tak</div><div>Hej Søren</div></body>", "Hej Søren", "x"), null);
});

test("text the user typed above the signature or quoted mail", () => {
  const owa = (own) => '<body><div class="elementToProof">' + own + '</div><div id="Signature"><div>Mvh Nis</div></div></body>';
  assert.equal(g.typedText(owa("<br>")), "");
  assert.equal(g.typedText(owa("Tak for sidst")), "Tak for sidst");
  assert.equal(g.typedText('<body><div><br></div><div id="x_Signature">Mvh</div></body>'), "");
  assert.equal(g.typedText('<body><p class=MsoNormal>&nbsp;</p><a name="_MailAutoSig">Mvh</a></body>'), "");
  assert.equal(g.typedText('<body><div><br></div><hr><div id="divRplyFwdMsg">Fra: Anne</div></body>'), "");
  assert.equal(g.typedText("<body><div><br></div></body>"), "");
  assert.equal(g.typedText("<body><div>Mvh Nis</div></body>"), "Mvh Nis", "unknown signature counts as typed");
});

test("greeting HTML copies the font of the first paragraph", () => {
  const owa =
    '<body><div class="elementToProof" style="font-family: Aptos, Calibri, sans-serif; font-size: 12pt; color: rgb(0, 0, 0); margin: 0px;"><br></div></body>';
  assert.equal(
    g.greetingHtml(owa, "Hej Anne"),
    '<div class="elementToProof" style="font-family: Aptos, Calibri, sans-serif; font-size: 12pt; color: rgb(0, 0, 0)">Hej Anne</div>' +
      '<div class="elementToProof" style="font-family: Aptos, Calibri, sans-serif; font-size: 12pt; color: rgb(0, 0, 0)"><br></div>'
  );
  const word = '<body lang=DA><div class=WordSection1><p class=MsoNormal><o:p>&nbsp;</o:p></p></div></body>';
  assert.equal(
    g.greetingHtml(word, "Hej <Anne>"),
    '<p class="MsoNormal">Hej &lt;Anne&gt;</p><p class="MsoNormal">&nbsp;</p>'
  );
  const signatureOnly = '<body><div>Med venlig hilsen</div><div style="font-size:8pt">Nis</div></body>';
  assert.equal(g.greetingHtml(signatureOnly, "Hej Anne"), "<div>Hej Anne</div><div><br></div>");
});
