const pptxgen = require("pptxgenjs");
const { applyTheme } = require("C:/Users/jaros/.claude/skills/synced/caaf1942-e05a-4076-bbc3-b00d061bdddc_04ccea6b-1963-4382-8ee4-ee30b52f392c/pptx/scripts/apply_theme.js");

const OUT = process.argv[2] || "Section03_Walking_Skeleton_React.pptx";

const THEME = {
  name: "Reactivities",
  headFontFace: "Calibri",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "20232A", lt1: "FFFFFF", dk2: "4A5568", lt2: "F1F5F9",
    accent1: "0B7EA3", accent2: "61DAFB", accent3: "512BD4",
    accent4: "C97A0A", accent5: "2E8B57", accent6: "C53030",
    hlink: "0B7EA3", folHlink: "512BD4",
  },
};
const DARK = THEME.colors.dk1;
const CODE_MUTED = "9AA7B8";
const MONO = "Courier New";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Section 3: Building a walking skeleton part 2 - React";
pres.author = "Building Internet Services";
const C = pres.SchemeColor;
const W = 13.333;
const FOOT = "Building Internet Services  ·  Section 3: Walking skeleton part 2 – React";

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: DARK },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.0, w: 7.6, h: 2.3, fontSize: 44, bold: true, color: C.background1, align: "left", valign: "bottom" }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.5, w: 7.6, h: 1.4, fontSize: 20, color: C.accent2, align: "left", valign: "top" }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "SECTION_DARK",
  background: { color: DARK },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.6, w: 8.0, h: 1.3, fontSize: 40, bold: true, color: C.background1, align: "left", valign: "top" }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.0, w: 8.0, h: 1.6, fontSize: 18, color: C.background2, align: "left", valign: "top" }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 12.1, h: 0.9, fontSize: 32, bold: true, color: C.text1, align: "left", valign: "middle" }, text: "" } },
    { text: { text: FOOT, options: { x: 0.6, y: 6.98, w: 9, h: 0.3, fontSize: 10, color: C.text2, margin: 0 } } },
  ],
  slideNumber: { x: 12.1, y: 6.98, w: 0.65, h: 0.3, fontSize: 10, color: C.text2, align: "right" },
});

// ---------- helpers ----------
let section = "";
function newSection(t) { section = t; pres.addSection({ title: t }); }
function content(titleText, notes) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(titleText, { placeholder: "title" });
  s.addNotes(notes.trim());
  return s;
}
function atom(s, cx, cy, r, color) {
  for (const rot of [0, 60, 120]) {
    s.addShape(pres.ShapeType.ellipse, { x: cx - r, y: cy - r * 0.38, w: 2 * r, h: r * 0.76, rotate: rot, line: { color, width: 2.25 }, fill: { type: "none" }, objectName: "Atom orbit " + rot });
  }
  const d = r * 0.22;
  s.addShape(pres.ShapeType.ellipse, { x: cx - d / 2, y: cy - d / 2, w: d, h: d, fill: { color }, line: { type: "none" }, objectName: "Atom nucleus" });
}
function divider(num, titleText, sub, notes) {
  const s = pres.addSlide({ masterName: "SECTION_DARK", sectionTitle: section });
  s.addText("PART " + num, { x: 0.8, y: 2.0, w: 4, h: 0.4, fontSize: 16, bold: true, color: C.accent2, charSpacing: 4, margin: 0, isTextBox: true });
  s.addText(titleText, { placeholder: "title" });
  s.addText(sub, { placeholder: "body" });
  atom(s, 10.6, 3.75, 1.7, C.accent2);
  s.addNotes(notes.trim());
  return s;
}
function code(s, o) {
  const fs = o.fs || 14;
  s.addShape(pres.ShapeType.roundRect, { x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: 0.08, fill: { color: C.text1 }, line: { type: "none" }, objectName: "Code panel" });
  let top = o.y + 0.12;
  if (o.file) {
    s.addText(o.file, { x: o.x + 0.25, y: o.y + 0.1, w: o.w - 0.5, h: 0.3, fontSize: 12, bold: true, color: C.accent2, fontFace: MONO, margin: 0, isTextBox: true });
    top = o.y + 0.45;
  }
  const runs = o.lines.map((ln, i) => {
    let color = C.background1, bold = false, text = ln;
    if (ln.startsWith("!")) { text = ln.slice(1); color = C.accent2; bold = true; }
    else if (/^\s*(\/\/|#)/.test(ln)) color = CODE_MUTED;
    return { text: text === "" ? " " : text, options: { color, bold, breakLine: i < o.lines.length - 1 } };
  });
  s.addText(runs, { x: o.x + 0.25, y: top, w: o.w - 0.4, h: o.y + o.h - top - 0.1, fontSize: fs, fontFace: MONO, valign: "top", margin: 0, isTextBox: true, objectName: "Code" });
}
function card(s, o) {
  s.addShape(pres.ShapeType.roundRect, { x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: 0.1, fill: { color: C.background2 }, line: { type: "none" }, objectName: "Card" });
  let tx = o.x + 0.3, tw = o.w - 0.6;
  if (o.badge) {
    s.addText(o.badge, { shape: pres.ShapeType.ellipse, x: o.x + 0.3, y: o.y + 0.28, w: 0.55, h: 0.55, fill: { color: o.color || C.accent1 }, color: C.background1, bold: true, fontSize: 16, align: "center", valign: "middle", margin: 0, objectName: "Badge" });
    tx = o.x + 1.05; tw = o.w - 1.3;
  }
  s.addText(o.head, { x: tx, y: o.y + 0.25, w: tw, h: 0.6, fontSize: o.headSize || 18, bold: true, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
  if (o.body) s.addText(o.body, { x: o.x + 0.3, y: o.y + 0.95, w: o.w - 0.6, h: o.h - 1.1, fontSize: o.fs || 14, color: C.text2, valign: "top", margin: 0, isTextBox: true, paraSpaceAfter: 6 });
}
function bullets(s, items, o) {
  const runs = items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1 } }));
  s.addText(runs, { x: o.x, y: o.y, w: o.w, h: o.h, fontSize: o.fs || 16, color: C.text1, valign: "top", margin: 0, paraSpaceAfter: 10, isTextBox: true });
}
function note(s, o) { // tinted call-out box
  const col = o.color || C.accent1;
  s.addShape(pres.ShapeType.roundRect, { x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: 0.08, fill: { color: C.background2 }, line: { color: col, width: 1.25 }, objectName: "Callout" });
  s.addText([{ text: o.label + "  ", options: { bold: true, color: col } }, { text: o.text, options: { color: C.text1 } }],
    { x: o.x + 0.25, y: o.y + 0.05, w: o.w - 0.5, h: o.h - 0.1, fontSize: o.fs || 14, valign: "middle", margin: 0, isTextBox: true });
}
function arrow(s, x1, y, x2, color) {
  s.addShape(pres.ShapeType.line, { x: x1, y, w: x2 - x1, h: 0, line: { color: color || C.text2, width: 2, endArrowType: "triangle" }, objectName: "Arrow" });
}
function box(s, o) {
  s.addText(o.text, { shape: pres.ShapeType.roundRect, rectRadius: 0.1, x: o.x, y: o.y, w: o.w, h: o.h, fill: { color: o.fill }, line: { type: "none" }, color: o.color || C.background1, fontSize: o.fs || 16, bold: true, align: "center", valign: "middle", margin: 4, objectName: "Box" });
}
function label(s, text, o) {
  s.addText(text, Object.assign({ fontSize: 14, color: C.text2, margin: 0, isTextBox: true, valign: "top" }, o));
}

// =====================================================================
// INTRO
// =====================================================================
newSection("Introduction");
{
  const s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: section });
  s.addText("Building a walking skeleton, part 2: React", { placeholder: "title" });
  s.addText("Section 3 · Building Internet Services\nASP.NET Core Web API + React", { placeholder: "body" });
  atom(s, 10.6, 3.75, 1.9, C.accent2);
  s.addNotes(`
OPENING (2 min)

SAY: Last time we built the back end of our walking skeleton: an ASP.NET Core API that reads activities from PostgreSQL and returns them as JSON. We tested it in Postman and Swagger. Today we build the other half: a React application that runs in the browser, calls that API and shows the activities on the screen.

SAY: "Walking skeleton" means the thinnest possible version of the application that still goes end to end: database, API, client. It does almost nothing, but every layer is connected. After today the skeleton can walk.

SAY: This class follows Section 3 of the course (11 short videos, about 1 hour 8 minutes). We will code it live together, so please open the repository and a terminal now.

CHECK BEFORE STARTING: everyone has Node.js installed (node -v) and the API from last week builds and runs.
`.trim());
}
{
  const s = content("Where we are: the client is the missing piece", `
THE BIG PICTURE (3 min)

SAY: Read the diagram from right to left. PostgreSQL stores the activities. The API, built last week, exposes them at /api/activities. Both of these are finished and tested.

SAY: The box on the left is today's work. It is a separate application with its own tooling (Node.js, npm, Vite), its own folder (client) and its own port (3000). The browser downloads it once, and then JavaScript running in the browser asks the API for data.

POINT OUT: The two applications only share one thing: the HTTP contract. The client needs to know the URL and the shape of the JSON. Nothing else. That is why we can write the client in TypeScript and the server in C#.

ASK: "What will the browser show if the API is not running?" (Answer: an empty list and an error in the console. The client still loads because it is a separate application.)

NOTE: Ports on the slide are the ones from our repository. In the course videos the API runs on https://localhost:5001.
`);
  const y = 2.0, h = 1.7;
  box(s, { x: 0.6, y, w: 3.4, h, fill: C.accent1, text: "React client\n(browser)", fs: 20 });
  box(s, { x: 4.95, y, w: 3.4, h, fill: C.accent3, text: "ASP.NET Core API\n(controllers)", fs: 20 });
  box(s, { x: 9.3, y, w: 3.4, h, fill: C.text2, text: "PostgreSQL\n(EF Core)", fs: 20 });
  arrow(s, 4.1, y + h / 2, 4.85);
  arrow(s, 8.45, y + h / 2, 9.2);
  label(s, "localhost:3000", { x: 0.6, y: y + h + 0.12, w: 3.4, h: 0.3, align: "center", fontFace: MONO, fontSize: 13 });
  label(s, "localhost:5057  /api/activities", { x: 4.95, y: y + h + 0.12, w: 3.4, h: 0.3, align: "center", fontFace: MONO, fontSize: 13 });
  label(s, "database: reactivities", { x: 9.3, y: y + h + 0.12, w: 3.4, h: 0.3, align: "center", fontFace: MONO, fontSize: 13 });
  s.addText("TODAY", { x: 0.6, y: 1.55, w: 3.4, h: 0.35, fontSize: 14, bold: true, color: C.accent1, align: "center", charSpacing: 3, margin: 0, isTextBox: true });
  s.addText("DONE IN SECTION 2", { x: 4.95, y: 1.55, w: 7.75, h: 0.35, fontSize: 14, bold: true, color: C.text2, align: "center", charSpacing: 3, margin: 0, isTextBox: true });
  s.addText("By the end of this class you can", { x: 0.6, y: 4.55, w: 12, h: 0.4, fontSize: 18, bold: true, color: C.text1, margin: 0, isTextBox: true });
  bullets(s, [
    "Create a React + TypeScript project with Vite and explain what each generated file does",
    "Load data from the API with useState and useEffect, and explain why CORS blocks the first attempt",
    "Describe the API data with a TypeScript type, style the page with Material UI, and switch to HTTPS and Axios",
  ], { x: 0.6, y: 5.05, w: 12.1, h: 1.7, fs: 16 });
}
{
  const s = content("Nine steps from an empty folder to a styled list", `
ROADMAP (2 min)

SAY: These are the nine steps of today, in the order of the course videos. The numbers in minutes are the video lengths, so you can find the right video when you revise at home.

SAY: I grouped them into three parts. Part A, steps 1 and 2: we create the project and look at what Vite gave us. Part B, steps 3 to 6: we connect the client to the API. This is the heart of the class, and it includes the famous CORS error. Part C, steps 7 to 9: we make it look better and make it more professional.

TIP FOR YOU: Steps 3 and 4 belong together. In step 3 the code is correct but the browser refuses to show the data. Do not fix it too quickly. Let students see the red error in the console first. They will meet it many times in their careers.

TIMING: Part A about 20 minutes, Part B about 40 minutes, Part C about 25 minutes.
`);
  const steps = [
    ["1", "Create the React project", "Vite, react-ts · 6 min", C.accent1],
    ["2", "Review the project files", "What Vite generated · 9 min", C.accent1],
    ["3", "Fetch data from the API", "useState, useEffect · 9 min", C.accent3],
    ["4", "Configure CORS on the API", "Why the browser blocks us · 7 min", C.accent3],
    ["5", "Activity type in TypeScript", "Typing the JSON · 8 min", C.accent3],
    ["6", "React developer tools", "Inspect components · 3 min", C.accent3],
    ["7", "Add Material UI", "Ready-made components · 9 min", C.accent5],
    ["8", "HTTPS for the client", "mkcert plugin · 6 min", C.accent5],
    ["9", "Axios for data fetching", "A nicer HTTP client · 6 min", C.accent5],
  ];
  const cw = 3.9, ch = 1.35, gx = 0.2, gy = 0.2, x0 = 0.6, y0 = 1.55;
  steps.forEach((st, i) => {
    const col = Math.floor(i / 3) === 0 ? i : i; // keep order row-wise
    const cx = x0 + (i % 3) * (cw + gx), cy = y0 + Math.floor(i / 3) * (ch + gy);
    s.addShape(pres.ShapeType.roundRect, { x: cx, y: cy, w: cw, h: ch, rectRadius: 0.1, fill: { color: C.background2 }, line: { type: "none" }, objectName: "Step card" });
    s.addText(st[0], { shape: pres.ShapeType.ellipse, x: cx + 0.25, y: cy + 0.38, w: 0.6, h: 0.6, fill: { color: st[3] }, color: C.background1, bold: true, fontSize: 18, align: "center", valign: "middle", margin: 0, objectName: "Step number" });
    s.addText(st[1], { x: cx + 1.05, y: cy + 0.25, w: cw - 1.2, h: 0.45, fontSize: 16, bold: true, color: C.text1, margin: 0, isTextBox: true, valign: "middle" });
    s.addText(st[2], { x: cx + 1.05, y: cy + 0.72, w: cw - 1.2, h: 0.4, fontSize: 13, color: C.text2, margin: 0, isTextBox: true, valign: "middle" });
  });
  const ly = 6.3;
  [["A  Create the client", C.accent1, 0.6], ["B  Talk to the API", C.accent3, 4.7], ["C  Polish", C.accent5, 8.8]].forEach(([t, col, x]) => {
    s.addShape(pres.ShapeType.ellipse, { x, y: ly + 0.07, w: 0.22, h: 0.22, fill: { color: col }, line: { type: "none" }, objectName: "Legend dot" });
    s.addText(t, { x: x + 0.35, y: ly, w: 3.5, h: 0.36, fontSize: 14, bold: true, color: C.text1, margin: 0, isTextBox: true, valign: "middle" });
  });
}

// =====================================================================
// PART A
// =====================================================================
newSection("A - Create the client");
divider("A", "Create the client", "A new React + TypeScript project with Vite, and a tour of the files it generates.", `
PART A (about 20 min in total)

SAY: We start with an empty folder and finish this part with a running React application at http://localhost:3000. No API calls yet, only the client.

DEMO PREP: Open a terminal in the repository root, next to the API, Application, Domain and Persistence folders. The client folder will be created there, as a sibling of API.

ASK before you start: "Who has used React before? Who has used any JavaScript framework?" Adjust how much time you spend on the React basics slide.
`);
{
  const s = content("Three tools, three jobs", `
WHY THESE TOOLS (3 min)

SAY: Three names appear today and students often mix them up, so let us separate them.

React is a library for building user interfaces. You describe what the screen should look like for the current data, and React updates the real page for you. It is only the UI part; it has no opinion about HTTP or routing.

TypeScript is JavaScript with types. The browser cannot run TypeScript, so it is compiled to JavaScript. The benefit is the same as in C#: the editor finds mistakes before you run the code, and autocomplete knows what properties an object has. You already think in types because of C#, so this will feel familiar.

Vite is the development tool. It creates the project, runs a development server with hot reload (you save the file and the browser updates in well under a second), and builds the optimised files for production.

ANALOGY for .NET students: React is like the Razor/Blazor component model, TypeScript is like the C# compiler, and Vite is like "dotnet new" plus "dotnet watch" plus "dotnet publish".
`);
  const cw = 3.9, y = 1.6, h = 3.6;
  card(s, { x: 0.6, y, w: cw, h, badge: "R", color: C.accent1, head: "React", body: "A library for building the user interface from components.\n\nYou describe the screen for the current data; React updates the page.", fs: 16 });
  card(s, { x: 4.7, y, w: cw, h, badge: "TS", color: C.accent3, head: "TypeScript", body: "JavaScript with static types, compiled to plain JavaScript.\n\nErrors show up in the editor, not in the browser at run time.", fs: 16 });
  card(s, { x: 8.8, y, w: cw, h, badge: "V", color: C.accent5, head: "Vite", body: "Project generator, development server and production bundler.\n\nSave a file and the browser updates almost instantly.", fs: 16 });
  note(s, { x: 0.6, y: 5.55, w: 12.1, h: 0.9, label: "For .NET developers:", text: "React is the component model, TypeScript plays the role of the C# compiler, and Vite is dotnet new + dotnet watch + dotnet publish." });
}
{
  const s = content("Step 1: create the project with one command", `
LIVE DEMO (6 min) - course video 18

DO: In the repository root run the first command. Vite may ask questions:
- "Ok to proceed?" to install create-vite: answer y.
- Framework: React. Variant: TypeScript. If the template flag is given, these questions are skipped.
- Newer versions may also ask about the React Compiler or about installing and starting immediately. Accept the defaults; they do not change what we do today.

SAY while it installs: npm install reads package.json and downloads the dependencies into node_modules. This folder is large and is never committed to git; Vite already created a .gitignore for it.

DO: Run npm run dev. Open the URL printed in the terminal. Show the default page with the counter. Click the counter.

DO: Open src/App.tsx, change some text, save. Point at the browser: it updated without a refresh. This is hot module replacement.

SAY: By default Vite uses port 5173. We change it to 3000 on the next slide, because our CORS configuration later will name the port explicitly.

COMMON PROBLEM: "npm is not recognized" means Node.js is not installed or the terminal was opened before installation. Restart the terminal.
`);
  code(s, { x: 0.6, y: 1.6, w: 7.3, h: 3.3, file: "terminal (repository root)", fs: 14, lines: [
    "# 1. generate the project in ./client",
    "!npm create vite@latest client -- --template react-ts",
    "",
    "# 2. download the dependencies",
    "cd client",
    "!npm install",
    "",
    "# 3. start the development server",
    "!npm run dev",
  ] });
  card(s, { x: 8.3, y: 1.6, w: 4.4, h: 3.3, head: "What you get", body: [
    { text: "A client folder next to API", options: { bullet: true, breakLine: true } },
    { text: "React and TypeScript configured", options: { bullet: true, breakLine: true } },
    { text: "A dev server with hot reload", options: { bullet: true, breakLine: true } },
    { text: "ESLint rules for React", options: { bullet: true } },
  ], fs: 15 });
  note(s, { x: 0.6, y: 5.25, w: 12.1, h: 0.9, label: "Requirement:", text: "Node.js must be installed. Check with node -v and npm -v before you start." });
}
{
  const s = content("Step 2: what Vite generated", `
PROJECT TOUR (5 min) - course video 19, first half

DO: Open the client folder in VS Code and walk through the files in this order.

index.html: SAY this is the only HTML page in the whole application. That is what "single-page application" means. It contains an empty div with id "root" and one script tag that loads main.tsx.

src/main.tsx: the entry point. It finds the root div and tells React to render our App component into it.

src/App.tsx: the root component. Everything we see on the screen starts here. Today we work almost only in this file.

package.json: SAY this is like the .csproj file. "dependencies" are the NuGet packages of the JavaScript world, and "scripts" are named commands: npm run dev, npm run build.

vite.config.ts: configuration of the dev server and the build. We edit it twice today.

tsconfig files: TypeScript compiler options. We do not touch them.

DO: Delete what we do not need, as in the video: the contents of App.css and the demo markup in App.tsx. Keep it minimal so students see only what matters.
`);
  code(s, { x: 0.6, y: 1.55, w: 5.2, h: 5.1, file: "client/", fs: 14, lines: [
    "!index.html",
    "package.json",
    "vite.config.ts",
    "tsconfig.json",
    "tsconfig.app.json",
    "tsconfig.node.json",
    "eslint.config.js",
    "node_modules/",
    "public/",
    "src/",
    "!  main.tsx",
    "!  App.tsx",
    "  App.css",
    "  index.css",
    "  assets/",
  ] });
  const rows = [
    ["index.html", "The single HTML page. Contains <div id=\"root\"> and loads main.tsx."],
    ["src/main.tsx", "Entry point. Renders the App component into the root div."],
    ["src/App.tsx", "The root component. Today we work mostly here."],
    ["package.json", "Dependencies and scripts. The .csproj of the client."],
    ["vite.config.ts", "Dev server and build settings (port, plugins)."],
  ];
  rows.forEach((r, i) => {
    const y = 1.55 + i * 1.02;
    s.addText(r[0], { x: 6.2, y, w: 6.5, h: 0.38, fontSize: 15, bold: true, color: C.accent1, fontFace: MONO, margin: 0, isTextBox: true, valign: "middle" });
    s.addText(r[1], { x: 6.2, y: y + 0.38, w: 6.5, h: 0.5, fontSize: 14, color: C.text1, margin: 0, isTextBox: true, valign: "top" });
  });
}
{
  const s = content("How the application starts in the browser", `
STARTUP FLOW (4 min) - course video 19, second half

SAY: Follow the arrows. The browser asks for the page and gets index.html. That file loads main.tsx. main.tsx calls createRoot on the root div and renders App. From this moment React controls everything inside that div.

EXPLAIN StrictMode: It is a development-only helper. It deliberately renders components twice and runs effects twice to help us find bugs. It does nothing in the production build. Remember this, because in ten minutes we will see our API request sent twice in the Network tab, and this is the reason.

EXPLAIN the exclamation mark: getElementById can return null, and TypeScript complains. The "!" tells TypeScript "trust me, this element exists".

DO: Edit vite.config.ts and add server: { port: 3000 }. Stop and restart npm run dev (configuration changes need a restart). Open http://localhost:3000.

SAY: We fix the port so that everybody in the room has the same address, and because the API will soon need to know exactly which origin is allowed to call it.
`);
  const y = 1.6;
  box(s, { x: 0.6, y, w: 2.6, h: 0.8, fill: C.text2, text: "index.html", fs: 15 });
  box(s, { x: 3.75, y, w: 2.6, h: 0.8, fill: C.accent1, text: "main.tsx", fs: 15 });
  box(s, { x: 6.9, y, w: 2.6, h: 0.8, fill: C.accent1, text: "<App />", fs: 15 });
  box(s, { x: 10.05, y, w: 2.65, h: 0.8, fill: C.accent5, text: "page on screen", fs: 15 });
  arrow(s, 3.25, y + 0.4, 3.7); arrow(s, 6.4, y + 0.4, 6.85); arrow(s, 9.55, y + 0.4, 10.0);
  code(s, { x: 0.6, y: 2.75, w: 6.6, h: 3.9, file: "src/main.tsx", lines: [
    "import { StrictMode } from 'react'",
    "import { createRoot } from 'react-dom/client'",
    "import './index.css'",
    "import App from './App.tsx'",
    "",
    "createRoot(document.getElementById('root')!)",
    "  .render(",
    "!    <StrictMode>",
    "!      <App />",
    "!    </StrictMode>,",
    "  )",
  ] });
  code(s, { x: 7.5, y: 2.75, w: 5.2, h: 2.75, file: "vite.config.ts", lines: [
    "export default defineConfig({",
    "!  server: {",
    "!    port: 3000",
    "!  },",
    "  plugins: [react()],",
    "})",
  ] });
  note(s, { x: 7.5, y: 5.75, w: 5.2, h: 0.9, label: "StrictMode", text: "runs effects twice in development only.", color: C.accent4 });
}
{
  const s = content("React in four ideas", `
REACT BASICS (5 min; shorten if the group knows React)

SAY: You need only four ideas to understand everything we write today.

1. Component. A component is a plain function whose name starts with a capital letter and which returns what should be on the screen. App is a component.

2. JSX. The thing the function returns looks like HTML but it is JavaScript. Inside curly braces you can write any JavaScript expression. Small differences from HTML: className instead of class, and every tag must be closed.

3. State. A component can remember values between renders with useState. It gives you the current value and a function to change it. When you call the setter, React calls your component function again and updates the page. Never change state directly; always use the setter.

4. Effect. Code that talks to the outside world (an HTTP request, a timer) goes into useEffect, so that it runs after the component is on the screen and not on every render.

ASK: "In the code on the right, what happens when I click the button?" (setCount is called, React re-runs Counter, the new number is shown.)

SAY: Functions whose names start with "use" are called hooks. They can only be called at the top level of a component.
`);
  const items = [
    ["1", "Component", "A function that returns what to show. Its name starts with a capital letter."],
    ["2", "JSX", "HTML-like syntax inside JavaScript. Use { } to insert any expression."],
    ["3", "State", "Data the component remembers. Changing it with the setter re-renders the component."],
    ["4", "Effect", "Code for the outside world (HTTP, timers). Runs after the render."],
  ];
  items.forEach((it, i) => {
    const y = 1.55 + i * 1.28;
    s.addText(it[0], { shape: pres.ShapeType.ellipse, x: 0.6, y: y + 0.1, w: 0.6, h: 0.6, fill: { color: C.accent1 }, color: C.background1, bold: true, fontSize: 18, align: "center", valign: "middle", margin: 0, objectName: "Idea number" });
    s.addText(it[1], { x: 1.45, y, w: 4.9, h: 0.4, fontSize: 18, bold: true, color: C.text1, margin: 0, isTextBox: true, valign: "middle" });
    s.addText(it[2], { x: 1.45, y: y + 0.42, w: 4.9, h: 0.75, fontSize: 14, color: C.text2, margin: 0, isTextBox: true, valign: "top" });
  });
  code(s, { x: 6.8, y: 1.55, w: 5.9, h: 5.1, file: "Counter.tsx (example)", lines: [
    "import { useState } from 'react'",
    "",
    "// 1. a component is a function",
    "function Counter() {",
    "  // 3. state: value + setter",
    "!  const [count, setCount] = useState(0)",
    "",
    "  // 2. JSX with { } expressions",
    "  return (",
    "    <button",
    "!      onClick={() => setCount(count + 1)}>",
    "      Clicked {count} times",
    "    </button>",
    "  )",
    "}",
  ] });
}

// =====================================================================
// PART B
// =====================================================================
newSection("B - Talk to the API");
divider("B", "Talk to the API", "Fetch the activities, meet the CORS error, fix it on the server, and give the data a type.", `
PART B (about 40 min in total)

SAY: Now we connect the two halves. This part has a planned failure in the middle: our first request will be blocked by the browser. That is intentional. Understanding why it is blocked is one of the most useful things you will learn in this course.

DEMO PREP: Start the API in a second terminal: dotnet run --project API --launch-profile http. Check in the browser that http://localhost:5057/api/activities returns JSON. Keep both terminals visible: API on one side, client on the other.
`);
{
  const s = content("Step 3: fetch the activities when the component loads", `
LIVE CODING (9 min) - course video 20

DO: Clear App.tsx and type this code step by step. Do not paste it.

Step 1 - state. SAY: We need somewhere to keep the activities. useState([]) creates state with an empty array as the starting value.

Step 2 - effect. SAY: We want to load data once, when the component appears. That is what useEffect with an empty dependency array does.

Step 3 - fetch. SAY: fetch is built into the browser. It returns a Promise. The first then converts the response body from JSON, the second stores the result in state. Storing it causes a re-render.

Step 4 - render. SAY: map turns every activity into a list item. React needs a unique "key" on every item in a list so it can track which item is which. We use the id from the database.

EXPECT: TypeScript underlines activity.id and activity.title in red, because useState([]) is inferred as an array of "never". SAY: we will fix this properly in step 5. In the video the instructor does the same thing.

DO: Save and look at the browser. The heading appears but the list is empty. Open DevTools (F12), Console tab. There is a red error. Go to the next slide.

NOTE: The URL uses our port 5057. The course video uses https://localhost:5001.
`);
  code(s, { x: 0.6, y: 1.55, w: 7.7, h: 5.15, file: "src/App.tsx", lines: [
    "import { useEffect, useState } from 'react'",
    "",
    "function App() {",
    "!  const [activities, setActivities] = useState([])",
    "",
    "!  useEffect(() => {",
    "    fetch('http://localhost:5057/api/activities')",
    "      .then(response => response.json())",
    "      .then(data => setActivities(data))",
    "!  }, [])",
    "",
    "  return (",
    "    <>",
    "      <h3>Reactivities</h3>",
    "      <ul>",
    "        {activities.map((activity) => (",
    "!          <li key={activity.id}>{activity.title}</li>",
    "        ))}",
    "      </ul>",
    "    </>",
    "  )",
    "}",
  ], fs: 12 });
  const pts = [
    ["State", "useState([]) holds the list. It starts empty."],
    ["Effect", "useEffect runs after the first render and starts the request."],
    ["[ ]", "Empty dependency array: run once, not after every render."],
    ["key", "Each list item needs a unique key so React can track it."],
  ];
  pts.forEach((p, i) => {
    const y = 1.55 + i * 1.3;
    s.addShape(pres.ShapeType.roundRect, { x: 8.6, y, w: 4.1, h: 1.15, rectRadius: 0.08, fill: { color: C.background2 }, line: { type: "none" }, objectName: "Point card" });
    s.addText(p[0], { x: 8.8, y: y + 0.1, w: 3.7, h: 0.35, fontSize: 15, bold: true, color: C.accent1, margin: 0, isTextBox: true, valign: "middle" });
    s.addText(p[1], { x: 8.8, y: y + 0.45, w: 3.7, h: 0.62, fontSize: 13, color: C.text1, margin: 0, isTextBox: true, valign: "top" });
  });
}
{
  const s = content("Why the dependency array matters", `
CONCEPT (4 min)

SAY: This slide explains the most common beginner bug in React.

Look at the loop on the left. If we fetch directly in the component body, or use useEffect without the second argument, this happens: the component renders, the fetch starts, the data arrives, we call setActivities, state changes, React renders again, the fetch starts again... forever. The API receives hundreds of requests per second.

The second argument of useEffect is the list of values the effect depends on. React runs the effect again only when one of those values changes. An empty array means "depends on nothing", so the effect runs once after the first render.

DEMO (optional, 1 minute, impressive): remove the [] and show the Network tab filling with requests. Put it back quickly.

SAY about StrictMode: With the correct code you will still see two requests in the Network tab in development. That is StrictMode from main.tsx: it mounts the component, unmounts it and mounts it again to check that our effect is safe. In production there is one request.

ASK: "When would we put a value into the array?" (For example an activity id: fetch the details again whenever the selected id changes.)
`);
  s.addText("Without [ ]: an endless loop", { x: 0.6, y: 1.55, w: 5.8, h: 0.4, fontSize: 18, bold: true, color: C.accent6, margin: 0, isTextBox: true });
  const lx = 0.6, bw = 2.6, bh = 0.75;
  box(s, { x: lx, y: 2.2, w: bw, h: bh, fill: C.text2, text: "render", fs: 15 });
  box(s, { x: lx + 3.2, y: 2.2, w: bw, h: bh, fill: C.text2, text: "fetch", fs: 15 });
  box(s, { x: lx + 3.2, y: 3.6, w: bw, h: bh, fill: C.text2, text: "setActivities", fs: 15 });
  box(s, { x: lx, y: 3.6, w: bw, h: bh, fill: C.accent6, text: "state changed", fs: 15 });
  arrow(s, lx + bw + 0.05, 2.2 + bh / 2, lx + 3.15);
  s.addShape(pres.ShapeType.line, { x: lx + 3.2 + bw / 2, y: 2.2 + bh + 0.05, w: 0, h: 3.6 - 2.2 - bh - 0.1, line: { color: C.text2, width: 2, endArrowType: "triangle" }, objectName: "Arrow down" });
  s.addShape(pres.ShapeType.line, { x: lx + bw + 0.05, y: 3.6 + bh / 2, w: 0.55, h: 0, line: { color: C.text2, width: 2, beginArrowType: "triangle" }, objectName: "Arrow left" });
  s.addShape(pres.ShapeType.line, { x: lx + bw / 2, y: 2.2 + bh + 0.05, w: 0, h: 3.6 - 2.2 - bh - 0.1, line: { color: C.accent6, width: 2, beginArrowType: "triangle" }, objectName: "Arrow up" });
  label(s, "Every state change triggers a new render, which starts a new request.", { x: 0.6, y: 4.65, w: 5.8, h: 0.7, fontSize: 14 });

  s.addText("The second argument decides when", { x: 7.0, y: 1.55, w: 5.7, h: 0.4, fontSize: 18, bold: true, color: C.accent1, margin: 0, isTextBox: true });
  s.addTable([
    [{ text: "Second argument", options: { bold: true, color: C.background1, fill: { color: C.text1 } } }, { text: "The effect runs", options: { bold: true, color: C.background1, fill: { color: C.text1 } } }],
    [{ text: "(missing)", options: { fontFace: MONO } }, "after every render"],
    [{ text: "[]", options: { fontFace: MONO, bold: true, color: C.accent1 } }, { text: "once, after the first render", options: { bold: true } }],
    [{ text: "[id]", options: { fontFace: MONO } }, "after the first render and whenever id changes"],
  ], { x: 7.0, y: 2.15, w: 5.7, colW: [2.1, 3.6], fontSize: 14, color: C.text1, border: { type: "solid", color: "D5DCE5", pt: 0.75 }, margin: [6, 8, 6, 8], rowH: 0.5 });
  note(s, { x: 0.6, y: 5.6, w: 12.1, h: 0.95, label: "Two requests in the Network tab?", text: "That is StrictMode in development. Production sends one.", color: C.accent4 });
}
{
  const s = content("The browser blocks the request: CORS", `
THE PLANNED FAILURE (4 min) - course video 21, first half

DO: Show the red error in the browser console and read it aloud slowly. The key words are "blocked by CORS policy" and "No 'Access-Control-Allow-Origin' header".

DO: Open the Network tab and click the request. SAY: Look, the status is 200. The API answered and sent the JSON. The API terminal shows the SQL query too. So the server worked. It is the browser that refuses to give the response to our JavaScript.

EXPLAIN: An origin is the combination of scheme, host and port. http://localhost:3000 and http://localhost:5057 have a different port, so they are different origins. By default browsers apply the same-origin policy: a script from one origin cannot read responses from another origin. This protects users: without it, any website you visit could read data from your bank using your cookies.

CORS, Cross-Origin Resource Sharing, is the way a server says "I allow this other origin to read my responses". It does this with a response header.

ASK: "Why did Postman work last week with the same URL?" (Postman is not a browser. The same-origin policy is enforced by browsers only. CORS is not server security; it is a browser protection for users.)
`);
  const y = 1.6;
  box(s, { x: 0.6, y, w: 3.6, h: 1.3, fill: C.accent1, text: "React client\nhttp://localhost:3000", fs: 15 });
  box(s, { x: 9.1, y, w: 3.6, h: 1.3, fill: C.accent3, text: "API\nhttp://localhost:5057", fs: 15 });
  arrow(s, 4.3, y + 0.4, 9.0);
  label(s, "GET /api/activities", { x: 4.3, y: y + 0.02, w: 4.7, h: 0.3, align: "center", fontFace: MONO, fontSize: 13 });
  s.addShape(pres.ShapeType.line, { x: 4.3, y: y + 0.95, w: 4.7, h: 0, line: { color: C.accent6, width: 2, beginArrowType: "triangle", dashType: "dash" }, objectName: "Blocked response" });
  label(s, "200 OK, but no permission header: blocked", { x: 4.3, y: y + 1.02, w: 4.7, h: 0.3, align: "center", fontSize: 13, color: C.accent6, bold: true });
  code(s, { x: 0.6, y: 3.25, w: 12.1, h: 1.35, file: "browser console", fs: 12, lines: [
    "!Access to fetch at 'http://localhost:5057/api/activities' from origin 'http://localhost:3000'",
    "!has been blocked by CORS policy: No 'Access-Control-Allow-Origin' header is present.",
  ] });
  card(s, { x: 0.6, y: 4.85, w: 5.9, h: 1.85, head: "Origin = scheme + host + port", body: "Port 3000 and port 5057 are different origins, so the same-origin policy applies.", fs: 14 });
  card(s, { x: 6.8, y: 4.85, w: 5.9, h: 1.85, head: "Who blocks it?", body: "The browser, not the API. That is why Postman and Swagger worked last week.", fs: 14 });
}
{
  const s = content("Step 4: allow the client origin in Program.cs", `
LIVE CODING (5 min) - course video 21, second half

DO: Open API/Program.cs. Add two things.

1. In the services part: builder.Services.AddCors(). This registers the CORS services in the dependency injection container.

2. In the pipeline part: app.UseCors(...). This adds the middleware that writes the Access-Control-Allow-Origin header to responses.

EXPLAIN the options: AllowAnyHeader and AllowAnyMethod mean we accept any HTTP header and any verb (GET, POST, PUT, DELETE). WithOrigins is the important one: only these origins are allowed. We list both http and https for port 3000 because we switch the client to HTTPS later today.

STRESS the order: middleware runs in the order it is written. UseCors must be before MapControllers, otherwise the response is produced before the header can be added. If we add authentication later, the order is UseCors, UseAuthentication, UseAuthorization, MapControllers.

DO: Restart the API (or let dotnet watch do it). Refresh the browser. The list of ten activities appears. This is the moment the skeleton walks: database to API to browser.

DO: In the Network tab, click the request and show the response header Access-Control-Allow-Origin: http://localhost:3000.

WARN: Do not use AllowAnyOrigin as a habit. It works, but in production you want to list exactly who may call your API.
`);
  code(s, { x: 0.6, y: 1.55, w: 7.6, h: 5.1, file: "API/Program.cs", lines: [
    "builder.Services.AddControllers();",
    "builder.Services.AddDbContext<AppDbContext>(...);",
    "!builder.Services.AddCors();",
    "",
    "var app = builder.Build();",
    "",
    "// Configure the HTTP request pipeline.",
    "!app.UseCors(x => x",
    "!    .AllowAnyHeader()",
    "!    .AllowAnyMethod()",
    "!    .WithOrigins(\"http://localhost:3000\",",
    "!                 \"https://localhost:3000\"));",
    "",
    "app.UseAuthorization();",
    "app.MapControllers();",
  ] });
  card(s, { x: 8.5, y: 1.55, w: 4.2, h: 1.55, badge: "1", head: "Register", body: "AddCors() adds the services.", fs: 14 });
  card(s, { x: 8.5, y: 3.3, w: 4.2, h: 1.55, badge: "2", head: "Use", body: "UseCors() adds the header.", fs: 14 });
  note(s, { x: 8.5, y: 5.05, w: 4.2, h: 1.6, label: "Order matters:", text: "UseCors must come before MapControllers.", color: C.accent4 });
}
{
  const s = content("Step 5: describe the JSON with a TypeScript type", `
LIVE CODING (5 min) - course video 22, first half

SAY: The list works, but TypeScript still shows red lines under activity.id and activity.title. TypeScript does not know what the API returns. We fix that by describing the shape of an activity.

DO: Show the C# entity on the left (Domain/Activity.cs) and build the TypeScript type on the right, property by property. Create the folders src/lib/types and the file index.d.ts.

POINT OUT the differences:
- C# properties are PascalCase, the JSON and TypeScript are camelCase. ASP.NET Core converts the names automatically when it serialises to JSON.
- double becomes number. TypeScript has only one number type.
- bool becomes boolean.
- DateTime becomes string. JSON has no date type, so the date travels as text in ISO format, for example 2026-11-25T10:00:00Z.

TIP: A fast way to create the type. Copy one object from the JSON response in Postman or the browser, and use a "JSON to TypeScript" converter. In the video the instructor does exactly this. Then check the result by hand.

SAY about .d.ts: a file ending in .d.ts is a declaration file. Types declared there are global in the project, so we do not need to import Activity in every component.

HONEST NOTE: this type is a copy of the C# class, kept in sync by hand. If we add a property on the server, we must add it here too.
`);
  code(s, { x: 0.6, y: 1.55, w: 5.9, h: 4.35, file: "Domain/Activity.cs  (C#)", lines: [
    "public class Activity",
    "{",
    "  public string Id { get; set; }",
    "  public required string Title { get; set; }",
    "!  public DateTime Date { get; set; }",
    "  public required string Description ...",
    "  public required string Category ...",
    "!  public bool IsCancelled { get; set; }",
    "  public required string City { get; set; }",
    "  public required string Venue { get; set; }",
    "!  public double Latitude { get; set; }",
    "!  public double Longitude { get; set; }",
    "}",
  ] });
  code(s, { x: 6.8, y: 1.55, w: 5.9, h: 4.35, file: "src/lib/types/index.d.ts  (TypeScript)", lines: [
    "type Activity = {",
    "  id: string",
    "  title: string",
    "!  date: string",
    "  description: string",
    "  category: string",
    "!  isCancelled: boolean",
    "  city: string",
    "  venue: string",
    "!  latitude: number",
    "!  longitude: number",
    "}",
  ] });
  note(s, { x: 0.6, y: 6.1, w: 12.1, h: 0.65, label: "Naming:", text: "C# uses PascalCase; the JSON and TypeScript use camelCase. ASP.NET Core converts the names for you." });
}
{
  const s = content("Using the type: errors move from the browser to the editor", `
LIVE CODING (3 min) - course video 22, second half

DO: Change one line in App.tsx: useState<Activity[]>([]). The angle brackets are a generic type argument, the same idea as List<Activity> in C#.

DO: Show that the red lines are gone. Then type "activity." inside the map and show the autocomplete list with all the properties.

DO: Make a deliberate mistake: write activity.tittle. The editor underlines it immediately, and npm run build would fail. SAY: without TypeScript this would be a silent bug: the page would show an empty list item and no error anywhere.

IMPORTANT LIMITATION, say it clearly: TypeScript types exist only at compile time. They are removed when the code is compiled to JavaScript. The type is our promise about what the API returns; nothing checks it at run time. If the API sends something different, TypeScript will not notice. Later in the course we add validation for data that comes from users.

ASK: "What is the type of the 'activity' parameter inside map now?" (Activity. TypeScript infers it from the array type; we do not need to write it.)
`);
  code(s, { x: 0.6, y: 1.55, w: 7.4, h: 2.3, file: "src/App.tsx", lines: [
    "// before: TypeScript infers never[]",
    "const [activities, setActivities] = useState([])",
    "",
    "// after: an array of Activity, starting empty",
    "!const [activities, setActivities] =",
    "!  useState<Activity[]>([])",
  ] });
  code(s, { x: 0.6, y: 4.1, w: 7.4, h: 2.55, file: "what the editor now tells you", lines: [
    "<li key={activity.id}>{activity.tittle}</li>",
    "",
    "!Property 'tittle' does not exist on type",
    "!'Activity'. Did you mean 'title'?",
  ] });
  card(s, { x: 8.3, y: 1.55, w: 4.4, h: 1.55, badge: "1", color: C.accent5, head: "Autocomplete", body: "The editor lists every property.", fs: 14 });
  card(s, { x: 8.3, y: 3.3, w: 4.4, h: 1.55, badge: "2", color: C.accent5, head: "Early errors", body: "Typos fail in the editor and build.", fs: 14 });
  note(s, { x: 8.3, y: 5.05, w: 4.4, h: 1.6, label: "Limit:", text: "types are removed at build time. Nothing checks the real API response.", color: C.accent4 });
}
{
  const s = content("Step 6: look inside the app with React Developer Tools", `
SHORT DEMO (3 min) - course video 23

DO: Install the "React Developer Tools" extension from the Chrome Web Store (it also exists for Firefox and Edge). Students can do it now; it takes thirty seconds. After installing, close and reopen DevTools.

DO: Open DevTools, find the new Components tab. Click App in the tree. On the right side, under "hooks", there is State with our array of ten activities. Expand one and show the properties.

SAY: This is the fastest way to answer the question "is my data wrong, or is my rendering wrong?". If the state here contains the activities but the page is empty, the problem is in the JSX. If the state is empty, the problem is in the request. Use the Network tab for that.

DO: Edit a title directly in the Components tab and show that the page changes. This is a nice way to show that the screen is a function of the state.

MENTION: The second tab, Profiler, records why and how often components render. We do not need it today.

DEBUGGING ROUTINE to give students: 1. Console tab for errors. 2. Network tab for the request and response. 3. Components tab for the state.
`);
  card(s, { x: 0.6, y: 1.55, w: 3.9, h: 2.5, badge: "1", head: "Console", body: "Red errors first: CORS, typos, failed requests.", fs: 14 });
  card(s, { x: 4.7, y: 1.55, w: 3.9, h: 2.5, badge: "2", head: "Network", body: "Was the request sent? What status and JSON came back?", fs: 14 });
  card(s, { x: 8.8, y: 1.55, w: 3.9, h: 2.5, badge: "3", color: C.accent3, head: "Components", body: "New tab from the extension: the component tree, props and state.", fs: 14 });
  code(s, { x: 0.6, y: 4.3, w: 7.4, h: 2.35, file: "Components tab  ›  App", lines: [
    "hooks",
    "!  State: [{…}, {…}, {…}, {…}, {…}, …]   // 10 items",
    "      0: { id: \"…\", title: \"Past Activity 1\",",
    "           category: \"drinks\", city: \"London\", … }",
    "  Effect: ƒ () {}",
  ] });
  note(s, { x: 8.3, y: 4.3, w: 4.4, h: 2.35, label: "Rule of thumb:", text: "state is correct but the page is wrong: look at the JSX. State is empty: look at the request." });
}

// =====================================================================
// PART C
// =====================================================================
newSection("C - Polish");
divider("C", "Polish", "Material UI components, HTTPS for the client, and Axios instead of fetch.", `
PART C (about 25 min in total)

SAY: The skeleton works. In this last part we make three improvements that the rest of the course depends on: a component library so we do not write CSS by hand, HTTPS in development, and a better HTTP client.

SAY: None of these changes what the application does. The same ten titles are shown at the end. But each one removes a problem we would meet later.
`);
{
  const s = content("Step 7: add Material UI", `
LIVE CODING (5 min) - course video 24, first half

SAY: Our list is plain HTML. We could write CSS ourselves, but most teams use a component library. This course uses Material UI (MUI), an implementation of Google's Material Design for React. It gives us ready-made components: buttons, cards, lists, dialogs, date pickers.

DO: Stop the dev server and run the install commands in the client folder.
- @mui/material is the library. @emotion/react and @emotion/styled are the styling engine it uses.
- @fontsource/roboto is the Roboto font, which Material Design expects. We install it from npm instead of loading it from Google, so the app works offline.
- @mui/icons-material is the icon set; we use it in later sections.

DO: Open src/main.tsx and add the four font imports. The numbers are font weights: light, regular, medium, bold.

DO: Delete the default styles in index.css (or the whole content), as in the video, so the Vite demo styles do not fight with MUI.

DO: Start npm run dev again. Nothing looks different yet, except the font. We use the components on the next slide.

SAY: Open package.json and show that the new packages were added to "dependencies". Anyone who clones the repository gets them with npm install.
`);
  code(s, { x: 0.6, y: 1.55, w: 7.5, h: 2.35, file: "terminal (client folder)", lines: [
    "# component library + its styling engine",
    "!npm install @mui/material @emotion/react @emotion/styled",
    "",
    "# Roboto font and the icon set",
    "!npm install @fontsource/roboto @mui/icons-material",
  ] });
  code(s, { x: 0.6, y: 4.15, w: 7.5, h: 2.5, file: "src/main.tsx", lines: [
    "import './index.css'",
    "!import '@fontsource/roboto/300.css'",
    "!import '@fontsource/roboto/400.css'",
    "!import '@fontsource/roboto/500.css'",
    "!import '@fontsource/roboto/700.css'",
    "import App from './App.tsx'",
  ] });
  card(s, { x: 8.4, y: 1.55, w: 4.3, h: 2.35, head: "Why a component library?", body: "Consistent, accessible, responsive components without writing CSS by hand.", fs: 14 });
  card(s, { x: 8.4, y: 4.15, w: 4.3, h: 2.5, head: "Why Roboto from npm?", body: "Material Design is built around Roboto. Installing it locally means no request to an external font server.", fs: 14 });
}
{
  const s = content("Replace HTML tags with MUI components", `
LIVE CODING (4 min) - course video 24, second half

DO: Change the JSX in App.tsx step by step: h3 becomes Typography, ul becomes List, li becomes ListItem with ListItemText inside.

EXPLAIN Typography: one component for all text. The "variant" prop chooses the style (h1 to h6, body1, caption...). Change h3 to h1 and back to show the effect.

EXPLAIN props: variant='h3' is a prop. Props are how a parent passes data to a child component, like parameters of a function. This is our first use of somebody else's components, and soon we will write our own with our own props.

POINT OUT: The key moved to ListItem, because key must be on the outermost element returned inside map.

SAY: The import line is important. VS Code can add it for you: put the cursor on the component name and press Ctrl + . (quick fix).

DO: Save and show the result. It is still a simple list, but now with Material typography and spacing. In the next sections this becomes cards with images, a navigation bar and forms.

DO: Open the MUI documentation (mui.com) and show the component list for two minutes. SAY: nobody remembers these components by heart. You find the one you need, copy the example, and adapt it.
`);
  code(s, { x: 0.6, y: 1.55, w: 5.4, h: 3.6, file: "before: plain HTML", lines: [
    "<>",
    "  <h3>Reactivities</h3>",
    "  <ul>",
    "    {activities.map((activity) => (",
    "      <li key={activity.id}>",
    "        {activity.title}",
    "      </li>",
    "    ))}",
    "  </ul>",
    "</>",
  ], fs: 13 });
  code(s, { x: 6.3, y: 1.55, w: 6.4, h: 3.6, file: "after: Material UI", lines: [
    "<>",
    "!  <Typography variant='h3'>Reactivities</Typography>",
    "!  <List>",
    "    {activities.map((activity) => (",
    "!      <ListItem key={activity.id}>",
    "!        <ListItemText>{activity.title}</ListItemText>",
    "!      </ListItem>",
    "    ))}",
    "!  </List>",
    "</>",
  ], fs: 13 });
  code(s, { x: 0.6, y: 5.4, w: 12.1, h: 1.25, file: "import (Ctrl + . adds it for you)", lines: [
    "import { List, ListItem, ListItemText, Typography } from '@mui/material'",
  ] });
}
{
  const s = content("Step 8: run the client over HTTPS", `
LIVE CODING (6 min) - course video 25

SAY: Why HTTPS on a development machine? Three reasons. First, production is always HTTPS, and we want development to be as close to production as possible. Second, later in the course we log in with cookies, and secure cookies need HTTPS. Third, a page served over HTTPS should call an API over HTTPS; mixing them causes "mixed content" problems in some browsers.

DO: Install the plugin as a development dependency (-D), because it is only needed while developing.

DO: Edit vite.config.ts: import mkcert and add mkcert() to the plugins array.

DO: Restart npm run dev. The first time, the plugin downloads the mkcert tool and creates a local certificate authority. Windows shows a security prompt asking to install the certificate; answer Yes. The address is now https://localhost:3000.

DO, for our repository: stop the API and start it with the https profile: dotnet run --project API --launch-profile https. The API is now at https://localhost:7161. Change the URL in App.tsx. If the browser does not trust the API certificate, run: dotnet dev-certs https --trust.

REMIND: This is why we wrote both http and https origins in the CORS configuration in step 4. https://localhost:3000 is a different origin from http://localhost:3000, because the scheme is part of the origin.

COMMON PROBLEM: on university computers without administrator rights the certificate cannot be installed. In that case stay on http for both client and API; everything else in the course still works.
`);
  code(s, { x: 0.6, y: 1.55, w: 6.3, h: 1.35, file: "terminal (client folder)", lines: [
    "!npm install -D vite-plugin-mkcert",
  ] });
  code(s, { x: 0.6, y: 3.15, w: 6.3, h: 3.5, file: "vite.config.ts", lines: [
    "import { defineConfig } from 'vite'",
    "import react from '@vitejs/plugin-react'",
    "!import mkcert from 'vite-plugin-mkcert'",
    "",
    "export default defineConfig({",
    "  server: {",
    "    port: 3000",
    "  },",
    "!  plugins: [react(), mkcert()],",
    "})",
  ] });
  s.addText("After the restart", { x: 7.3, y: 1.55, w: 5.4, h: 0.4, fontSize: 18, bold: true, color: C.text1, margin: 0, isTextBox: true });
  box(s, { x: 7.3, y: 2.1, w: 5.4, h: 0.8, fill: C.accent1, text: "client  https://localhost:3000", fs: 15 });
  box(s, { x: 7.3, y: 3.15, w: 5.4, h: 0.8, fill: C.accent3, text: "API  https://localhost:7161", fs: 15 });
  bullets(s, [
    "Same as production, where everything is HTTPS",
    "Needed later for secure login cookies",
    "No mixed http and https requests",
  ], { x: 7.3, y: 4.25, w: 5.4, h: 1.5, fs: 14 });
  note(s, { x: 7.3, y: 5.75, w: 5.4, h: 0.9, label: "API certificate:", text: "dotnet dev-certs https --trust", color: C.accent4 });
}
{
  const s = content("Step 9: replace fetch with Axios", `
LIVE CODING (6 min) - course video 26

SAY: fetch works, so why change? Axios is a small library that removes repetitive code and adds features we need later.

DO: npm install axios. Then change the effect in App.tsx.

WALK THROUGH the differences on the slide:
1. No response.json() step. Axios parses JSON automatically and puts the result in response.data.
2. It is typed. axios.get<Activity[]> tells TypeScript what response.data is. With fetch, json() returns "any" and we lose the type.
3. Errors. fetch only rejects when the network fails. A 404 or 500 response is a success for fetch, and you must check response.ok yourself. Axios rejects the promise for every status outside 200 to 299, so one catch handles all errors.
4. Interceptors. Later we configure one Axios instance with the base URL, and add code that runs for every request and response: show a loading indicator, show a toast on errors, send the login cookie. With fetch we would write that ourselves.

ABOUT "return () => {}": the function returned from an effect is the cleanup function. React calls it when the component is removed. Here it is empty; the instructor adds it as a reminder. It is optional.

DO: Save and show that the page looks exactly the same. Open the Network tab: same request, same response. We changed how we ask, not what we ask.
`);
  code(s, { x: 0.6, y: 1.55, w: 6.5, h: 3.2, file: "src/App.tsx", lines: [
    "!import axios from 'axios'",
    "",
    "useEffect(() => {",
    "!  axios.get<Activity[]>(",
    "!      'https://localhost:7161/api/activities')",
    "!    .then(response => setActivities(response.data))",
    "",
    "  return () => {}",
    "}, [])",
  ], fs: 13 });
  code(s, { x: 0.6, y: 5.0, w: 6.5, h: 1.3, file: "terminal (client folder)", lines: ["!npm install axios"] });
  const hdr = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text1 } } });
  s.addTable([
    [hdr(""), hdr("fetch"), hdr("Axios")],
    [{ text: "JSON", options: { bold: true } }, "call response.json()", "parsed: response.data"],
    [{ text: "Types", options: { bold: true } }, "result is any", "get<Activity[]>()"],
    [{ text: "HTTP 404 / 500", options: { bold: true } }, "not an error; check response.ok", "rejects: one catch"],
    [{ text: "Shared setup", options: { bold: true } }, "write it yourself", "base URL, interceptors"],
    [{ text: "Install", options: { bold: true } }, "built into the browser", "npm install axios"],
  ], { x: 7.4, y: 1.55, w: 5.3, colW: [1.5, 1.9, 1.9], fontSize: 13, color: C.text1, border: { type: "solid", color: "D5DCE5", pt: 0.75 }, margin: [5, 6, 5, 6], rowH: 0.62 });
}

// =====================================================================
// WRAP-UP
// =====================================================================
newSection("Wrap-up");
{
  const s = content("The finished App.tsx", `
REVIEW (3 min)

SAY: This is the complete file at the end of Section 3. About thirty lines. Let us read it from top to bottom once more, because every line was a decision we made today.

- Imports: MUI components, axios, and two hooks from React.
- State: a typed array of activities, empty at the start.
- Effect: runs once after the first render, asks the API, stores the result. Storing it re-renders the component.
- JSX: a heading and a list. map creates one ListItem per activity, each with a key.

ASK the group to trace the timeline: "What does the user see in the first 100 milliseconds?" (The heading and an empty list, because the state is still an empty array.) "And then?" (The response arrives, setActivities is called, React renders again with ten items.)

SAY: There is no loading indicator and no error handling yet. If the API is down, the user sees an empty list. That is acceptable for a walking skeleton, and we improve it later in the course.

DO: Commit the work. git add, git commit -m "End of section 3". Check that node_modules is not in the commit.
`);
  code(s, { x: 0.6, y: 1.5, w: 8.3, h: 5.3, file: "src/App.tsx", fs: 11, lines: [
    "import { List, ListItem, ListItemText, Typography } from '@mui/material'",
    "import axios from 'axios'",
    "import { useEffect, useState } from 'react'",
    "",
    "function App() {",
    "!  const [activities, setActivities] = useState<Activity[]>([])",
    "",
    "!  useEffect(() => {",
    "    axios.get<Activity[]>('https://localhost:7161/api/activities')",
    "      .then(response => setActivities(response.data))",
    "!  }, [])",
    "",
    "  return (",
    "    <>",
    "      <Typography variant='h3'>Reactivities</Typography>",
    "      <List>",
    "        {activities.map((activity) => (",
    "!          <ListItem key={activity.id}>",
    "            <ListItemText>{activity.title}</ListItemText>",
    "          </ListItem>",
    "        ))}",
    "      </List>",
    "    </>",
    "  )",
    "}",
  ] });
  const tl = [["1", "First render", "Heading and an empty list."], ["2", "Effect runs", "Axios sends GET /api/activities."], ["3", "Data arrives", "setActivities stores ten items."], ["4", "Second render", "The list shows ten titles."]];
  tl.forEach((t, i) => {
    const y = 1.5 + i * 1.32;
    s.addText(t[0], { shape: pres.ShapeType.ellipse, x: 9.2, y: y + 0.08, w: 0.55, h: 0.55, fill: { color: C.accent1 }, color: C.background1, bold: true, fontSize: 16, align: "center", valign: "middle", margin: 0, objectName: "Timeline number" });
    s.addText(t[1], { x: 9.95, y, w: 2.75, h: 0.4, fontSize: 16, bold: true, color: C.text1, margin: 0, isTextBox: true, valign: "middle" });
    s.addText(t[2], { x: 9.95, y: y + 0.42, w: 2.75, h: 0.7, fontSize: 13, color: C.text2, margin: 0, isTextBox: true, valign: "top" });
  });
}
{
  const s = content("When it does not work: five usual suspects", `
TROUBLESHOOTING (use during the lab, or 3 min as a summary)

SAY: Keep this slide open during the exercise. Almost every problem in today's lab is one of these five.

1. CORS error. Check three things: AddCors is registered, UseCors is before MapControllers, and the origin in WithOrigins matches the address bar exactly, including http or https and without a slash at the end. Then restart the API.

2. ERR_CONNECTION_REFUSED. Nothing is listening at that address. The API is not running, or the port in the URL is wrong. Compare with the "Now listening on" line in the API terminal.

3. ERR_CERT_AUTHORITY_INVALID. The browser does not trust the development certificate of the API. Run dotnet dev-certs https --trust and restart the browser.

4. Empty list, no error. Use the routine: Network tab first. Is there a request? Is the response an array with items? If the response is an empty array, the database is empty: check the seeding from Section 2 and that PostgreSQL in WSL is running.

5. The "key" warning. The key prop is missing or it is on the wrong element. It must be on the outermost element inside map.

TEACHING TIP: when a student calls you, do not look at the code first. Ask them: "What does the Console say? What does the Network tab say?" It builds the habit.
`);
  const hdr = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text1 } } });
  const m = (t) => ({ text: t, options: { fontFace: MONO, fontSize: 12, bold: true, color: C.accent6 } });
  s.addTable([
    [hdr("What you see"), hdr("Likely cause"), hdr("Fix")],
    [m("blocked by CORS policy"), "Origin not allowed, or UseCors is too late in the pipeline", "Check WithOrigins and the order; restart the API"],
    [m("ERR_CONNECTION_REFUSED"), "API not running, or wrong port in the URL", "Start the API; compare with \"Now listening on\""],
    [m("ERR_CERT_AUTHORITY_INVALID"), "API development certificate is not trusted", "dotnet dev-certs https --trust"],
    [{ text: "Empty list, no error", options: { bold: true } }, "The API returned an empty array", "Check the Network tab, the seed data and PostgreSQL"],
    [{ text: "Warning: unique \"key\" prop", options: { bold: true } }, "key is missing or on an inner element", "Put key on the outermost element inside map"],
  ], { x: 0.6, y: 1.6, w: 12.1, colW: [3.7, 4.2, 4.2], fontSize: 14, color: C.text1, border: { type: "solid", color: "D5DCE5", pt: 0.75 }, margin: [6, 8, 6, 8], rowH: 0.78, valign: "middle" });
}
{
  const s = content("Practice: make the skeleton your own", `
LAB EXERCISES (remaining time, or homework)

SAY: Three exercises, from easy to harder. All of them use only what we did today.

Exercise 1 (5 minutes). Show more data. ListItemText has a "secondary" prop for a second line of text. The goal is to practise reading properties from the typed object and to see autocomplete working. Possible solution: <ListItemText primary={activity.title} secondary={activity.city + ', ' + activity.venue} />.

Exercise 2 (10 minutes). Break CORS on purpose. Students change the port in WithOrigins to 3001, restart the API, and describe exactly what they see in the Console and in the Network tab. Then they fix it. The goal is to recognise the error quickly in the future. Ask one student to explain to the group why the status code is still 200.

Exercise 3 (15 minutes). The details endpoint. Last week we wrote GET /api/activities/{id}. Students add a second piece of state for one Activity and load it with axios.get<Activity>. They can hard-code an id copied from the list response. Stronger students can load it when a list item is clicked (onClick on ListItem). This prepares the ground for the details view in the next sections.

CHECK at the end: everyone has committed and pushed, and node_modules is not in the repository.
`);
  const ex = [
    ["1", "Show more data", "Add the city and venue as a second line under each title.", "Hint: the secondary prop of ListItemText", C.accent5],
    ["2", "Break CORS on purpose", "Change the allowed origin to port 3001. Describe what the Console and Network tabs show, then fix it.", "Question: why is the status still 200?", C.accent4],
    ["3", "Load one activity", "Call GET /api/activities/{id} with Axios and show the description of one activity.", "Hint: useState<Activity>() and axios.get<Activity>", C.accent3],
  ];
  ex.forEach((e, i) => {
    const x = 0.6 + i * 4.1, w = 3.9, y = 1.6, h = 4.6;
    s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.1, fill: { color: C.background2 }, line: { type: "none" }, objectName: "Exercise card" });
    s.addText(e[0], { shape: pres.ShapeType.ellipse, x: x + 0.3, y: y + 0.3, w: 0.7, h: 0.7, fill: { color: e[4] }, color: C.background1, bold: true, fontSize: 22, align: "center", valign: "middle", margin: 0, objectName: "Exercise number" });
    s.addText(e[1], { x: x + 0.3, y: y + 1.2, w: w - 0.6, h: 0.5, fontSize: 18, bold: true, color: C.text1, margin: 0, isTextBox: true, valign: "middle" });
    s.addText(e[2], { x: x + 0.3, y: y + 1.8, w: w - 0.6, h: 1.5, fontSize: 14, color: C.text1, margin: 0, isTextBox: true, valign: "top" });
    s.addText(e[3], { x: x + 0.3, y: y + 3.4, w: w - 0.6, h: 0.9, fontSize: 13, italic: true, color: C.text2, margin: 0, isTextBox: true, valign: "top" });
  });
}
{
  const s = pres.addSlide({ masterName: "SECTION_DARK", sectionTitle: section });
  s.addText("SUMMARY", { x: 0.8, y: 0.8, w: 4, h: 0.4, fontSize: 16, bold: true, color: C.accent2, charSpacing: 4, margin: 0, isTextBox: true });
  s.addText("The skeleton walks: database to API to browser", { x: 0.8, y: 1.3, w: 8.2, h: 1.5, fontSize: 36, bold: true, color: C.background1, margin: 0, isTextBox: true, valign: "top" });
  const pts = [
    "Vite creates and runs a React + TypeScript client on port 3000",
    "useState keeps data; useEffect with [ ] loads it once",
    "CORS is enforced by the browser and allowed by the server",
    "A TypeScript type describes the JSON the API returns",
    "Material UI, HTTPS and Axios prepare us for the next sections",
  ];
  s.addText(pts.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < pts.length - 1 } })), { x: 0.8, y: 3.0, w: 8.2, h: 3.0, fontSize: 18, color: C.background1, margin: 0, isTextBox: true, paraSpaceAfter: 10, valign: "top" });
  s.addText("Next: back to the API and the Application layer", { x: 0.8, y: 6.3, w: 8.2, h: 0.5, fontSize: 16, italic: true, color: C.accent2, margin: 0, isTextBox: true });
  atom(s, 10.9, 3.75, 1.6, C.accent2);
  s.addNotes(`
SUMMARY (3 min) - course video 27

SAY: Let us repeat the five things to remember from today.

1. Vite gave us a React and TypeScript project with one command. The client is a separate application in the client folder, with its own port.

2. Two hooks do all the work: useState keeps the data, useEffect with an empty dependency array loads it once.

3. CORS. The browser protects users with the same-origin policy. Our API allows our client explicitly with UseCors. Remember: it is the browser that blocks, and the server that gives permission.

4. The Activity type gives us autocomplete and compile-time checking, but it is a promise, not a run-time check.

5. Material UI, HTTPS and Axios did not change what the user sees, but everything in the following sections builds on them.

SAY: We now have a complete walking skeleton. Every layer exists and they are connected. From now on we add features through all the layers at once. Next time we go back to the server side and start filling the Application project, which is still empty, with real use cases: create, edit and delete activities.

HOMEWORK: watch videos 17 to 27 of the course (1 hour 8 minutes) and finish the three exercises.

QUESTIONS: leave two or three minutes.
`.trim());
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();
