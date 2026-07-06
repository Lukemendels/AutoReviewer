# StickShift Tool File — The Dual-Life HTML Pattern

First-principles specification — 2026-06-21. Extends the StickShift Platform
Spec and OKF Orchestration V2 (§A2, §A5, §A7). Defines the *single artifact*
that is both a standalone human utility and an agent-orchestrable tool — with
no fork, no lite version, no drift.

**How to read this.** Part I is the durable reasoning. Part II is the file
anatomy that falls out of it — including the quiet affordance and the
*Welcome to StickShift* screen. Part III is what it preserves and costs.
Part IV is the build ledger. If only Part I survives, the file is
re-derivable.

---

## Part I — First Principles

### P1. One file, two lives — never a fork

The standalone tool and the agentic tool are not two builds. They are one
`.html`, read two ways. The double-click human gets a sealed rung-2 utility
that manufactures a deliverable and hands it back. The agent gets a tool whose
embedded contract it can orchestrate. **Nothing is added or removed between the
two lives** — the second life is *latent cargo* that costs the first life
nothing.

The instant you maintain a "lite version" and a "full version," drift sets in
and the dual-life property dies. The rule: **ship one file that knows how to be
used two ways.**

### P2. The agent is just another supplier

This is the Economic Modeler's slot pattern (V2 §P7, §A4), one rung up. There,
the tool had one input slot and three suppliers behind it — Power Automate, a
VBA fetch, a human with a CSV — and the tool never learned which one filled the
slot.

Here, the **agent is a fourth supplier**:

- **Standalone life:** the *human* fills the slot — typing, pasting, uploading
  through the UI.
- **Agentic life:** the *agent* fills the slot — the model emits a `call:`
  payload the human carries in and pastes.

Same slot, same compute, same output. The tool is **agent-agnostic by
construction** — it consumes what is in the slot and does not care who put it
there. That indifference is *why* there is no fork: "the version without an
agent" is not a stripped build, it is the full tool with the slot filled by
hand.

### P3. Two readers, two surfaces

A human and an agent read different things in the same file, and you must not
force either to read the other's surface:

| Reader | Surface | Carrier |
|---|---|---|
| Standalone human | the **UI** — labels, slot, Run, Download | good copy |
| Agent | the embedded **`skill.md`** contract | a string in the file |

This is StickShift's human/agent-interoperability claim, localized to one
artifact. Do not try to make the human read the `skill.md` — the UI carries
the human. Do not surface the `skill.md` by default — it is machine-facing
cargo. **The human is served by the screen; the agent is served by the
contract; the file holds both.**

### P4. The asymmetry — unready not nagged, ready not blocked

The whole "not obvious, then obvious" trick is a single asymmetry:

- The person who **never** touches the agent must never feel they are using a
  crippled demo. The standalone tool earns its keep *fully*. No nag, no teaser,
  no "unlock the real version."
- The person who **is** ready must find the agentic capability waiting — in two
  seconds, in operator language, without having to be told it exists.

A standalone tool that feels like a locked trial breaks the trust that makes
the ramp work. A latent capability that can't be found is dead cargo. The
design target is the narrow band that serves both: **complete on first read,
discoverable on second.**

---

## Part II — Architecture (derived from Part I)

### A1. File anatomy — the regions of one `.html`

A single self-contained, offline file with four regions:

```
┌─ stickshift-tool.html ───────────────────────────────┐
│  1. TOOL UI         the human's surface (P3)          │
│       input slot · Run · Download                     │
│                                                       │
│  2. COMPUTE         vendored inline; no network (P3,  │
│       transform / manufacture        StickShift §net) │
│                                                       │
│  3. EMBEDDED CONTRACT   skill.md as a string (agent   │
│       + register envelope            surface, P3)     │
│                                                       │
│  4. QUIET AFFORDANCE → WELCOME SCREEN  the bridge     │
│       (A3)                            (P4)            │
└───────────────────────────────────────────────────────┘
```

Regions 1–2 are the standalone life. Region 3 is dormant until invoked. Region
4 is the only place the two lives touch the same pixel.

### A2. The input slot accepts either fill (P2)

The slot is the tool's sole interface. It accepts:

- **human fill** — pasted text, typed fields, or an uploaded file (still
  sandboxed); and
- **agent fill** — a pasted `call:` payload in the tool's documented grammar.

The tool parses what it gets and runs. It does not branch on *who* filled it.
(If a tool's `call:` grammar and its human-entry fields differ in shape, the
parser normalizes both to the same internal input — the slot is the
normalization point, not two code paths.)

### A3. The quiet affordance and the *Welcome to StickShift* screen

The default view is a clean tool — slot, Run, Download — with exactly **one**
low-weight agentic affordance and no other chrome:

```
  ┌─────────────────────────────────────────────┐
  │  Paste your draft                             │
  │  [ .......................................... ]│
  │                                               │
  │            [ Run ]   [ Download .docx ]        │
  │                                               │
  │  ───────────────────────────────────────────  │
  │  ⚙ Part of StickShift  ▸                       │  ← the stumble
  └─────────────────────────────────────────────┘
```

Click it and you do **not** get raw markdown dumped on screen. You get the
*Welcome to StickShift* screen — **reward the stumble, greet the initiated:**

```
  ┌─────────────────────────────────────────────┐
  │  Welcome to StickShift                        │
  │                                               │
  │  This tool works on its own — you just used   │
  │  it. It is also one station in a governed AI  │
  │  workflow on the approved stack. If you run    │
  │  an OKF bundle, this tool can register itself  │
  │  for your agent to drive.                     │
  │                                               │
  │  ▸ Agent contract (skill.md)      [ Copy ]    │
  │  ▸ Register this tool (<VBA_WRITE>)[ Copy ]    │
  │                                               │
  │  New here? → what StickShift is (one screen)  │
  │                          [ ← back to tool ]    │
  └─────────────────────────────────────────────┘
```

The screen does **double duty by audience** (P4):

- **The stumbler** (curious, uninitiated) gets an orientation and an
  invitation — a discovery moment that converts curiosity into awareness
  *after* they already got value, never as a gate before it.
- **The initiated** gets the two artifacts they came for — the `skill.md` and
  the register envelope — copyable, in two seconds, then `back to tool`.

Sutton's eye never travels here; she sees a tool that reviews her redlines.
Paul finds it instantly. Neither is nagged or blocked.

> **Default-state rule:** before first click, the file shows zero agentic
> instruction beyond the single quiet line. The Welcome screen is *pulled*,
> never *pushed*.

### A4. Self-registration (V2 §A5, generalized to HTML)

The file carries its own `skill.md` as a string and, on demand, emits a
`<VBA_WRITE>` packet for it — **self-registration with no hop to the LLM,**
both halves data-class so nothing waits on a code import. Two triggers, same
mechanism:

- **Operator-driven:** the `Copy` next to *Register this tool* on the Welcome
  screen yields the `<VBA_WRITE>` envelope; the operator applies it via the
  dashboard's existing Apply-Write path.
- **First-run-in-bundle (optional):** if the tool detects it is being run from
  a bundle context, it can emit the same envelope automatically for the
  operator to apply.

Either way the emitted contract is **`provenance: generated:html:<name>`** —
machine-owned, re-emitted on next run, write-gate-protected. The linter's
`html`-path check (V2 §A7) backstops a tool whose file went missing.

### A5. `returns:` is unchanged — only the consumer differs (V2 §A2)

The tool emits the same output in both lives; who consumes it is the only
difference:

| `returns:` | Standalone human | Agent-driven |
|---|---|---|
| `file-terminal` | downloads the `.docx`; done | model says "you'll get a file, we're done" |
| `markdown-store` | copies the markdown out for their own use | model repacks → Apply Write stores it |
| `markdown-continue` | reads the value | model reasons on with it |
| `file-receipt` | downloads + sees the receipt | model logs the receipt |

The standalone human is *always* effectively at a terminal — they place the
output themselves. The agent path is what gives an output a trace back in the
base. Same emission, two destinations.

### A6. Inherited constraints (non-negotiable)

Nothing in the dual-life pattern relaxes StickShift's gates:

- **Offline, self-contained.** No CDN, no telemetry, libraries vendored inline.
  No-network is a hard requirement, not a nicety — it is also what lets the file
  hand to any operator on a locked workstation with zero install.
- **Sealed sandbox.** The file cannot reach the filesystem, clipboard (unless
  pasted in), or host Office apps. Outputs are offered as a manual OS download.
- **Data-class.** Because both lives stay inside the host and network
  boundaries, the file is reviewable-by-inspection — *the agentic cargo adds no
  review surface,* because emitting a markdown string and offering a download
  are both data-class acts.

---

## Part III — What it preserves, and what it costs

**Preserves.** Human-in-the-loop at every junction — the agent fills a slot a
human still carries and pastes; the human still clicks Run and places every
output. Adding the agentic life inserts no autonomous path. Data-class status
is intact; the dual-life file is no harder to approve than a single-life one.

**Buys.** A zero-cost adoption ramp. The tool delivers value at double-click
with no platform to learn; the platform becomes a *discovered upgrade*, pulled
by the ready, invisible to the rest. The captive audience converts itself,
later, because the upgrade was waiting inside the tool they already trusted.

**Costs.** One real discipline and one small risk:

- **Discipline:** the standalone life must be *genuinely complete*. The moment
  it feels like a teaser, P4 breaks and the ramp inverts into a nag.
- **Risk:** maintaining the embedded `skill.md` against the tool's actual
  `call:` grammar. If they drift, the agent orchestrates a contract the tool no
  longer honors. The linter's drift check (V2 §A7) is the backstop; until then
  it is a review habit.

---

## Part IV — Build ledger (YAGNI)

**Build first — AutoReviewer as the proof.** It is already the V2 first cut.
Make it the first dual-life file:

1. **Standalone life complete** — slot accepts a draft, Run produces a
   tracked-change `.docx`, Download delivers it. This is the version handed to
   Sutton: one file, double-click, five-minute first win, no agent.
2. **Embedded `skill.md` string** — the agent contract, dormant.
3. **The quiet affordance → Welcome screen** — single footer line; click yields
   the greeting + `skill.md` copy + register-envelope copy.
4. **Self-register on demand** — the `<VBA_WRITE>` emission from the Welcome
   screen (V2 §A5).

**Define now, build at trigger:**

- *First-run-in-bundle auto-emit* (A4) → build when the first bundle actually
  consumes the tool; until then the operator-driven copy covers it.
- *Slot normalization for divergent grammars* (A2) → only if a tool's human
  fields and `call:` payload genuinely differ in shape; AutoReviewer may not
  need it.
- *Shared Welcome-screen component* → at the **second** dual-life tool, which is
  what reveals which parts are truly common. The first tool can inline it.

**Open decisions to settle before the relevant piece:**

1. **Affordance label.** "Part of StickShift," "Agent integration," or a gear
   glyph alone — tested against not drawing the standalone user's eye.
2. **Welcome-screen depth.** Does *New here?* link to a one-screen explainer
   bundled in-file (offline, consistent) or stay a single paragraph? Leaning
   in-file: a stumble that hits a dead link is a worse first impression than no
   link.
3. **Auto-emit detection.** How does the file know it is "in a bundle" without
   network or host access? Likely a URL flag the bundle's link carries
   (`tool.html#bundle`) — decided per the actual orchestration link.

Nothing here changes a deployed tool until a task is cut from "build first."
The honest first task is **AutoReviewer, dual-life** — the smallest artifact
that proves standalone value, latent capability, and the stumble-reward bridge
at once, with a real skeptic waiting on the other side.
