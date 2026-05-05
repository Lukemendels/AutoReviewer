__AI CAPABILITIES__

__LOOK BOOK__

Transportation Security Administration

Office of Policy, Plans & Engagement  |  Strategy, Plans & Innovation

April 2026

# __Executive Summary__

TSA already has a complete AI builder stack\. No new budget\. No new systems\. No new approvals\.

This look book documents a portfolio of AI\-powered workflows built and deployed inside TSA’s existing M365 and DHSChat environment\. Every tool described here is real, operational, and built under the same constraints every analyst faces: no Python, no external infrastructure, limited tooling\.

__18,800__

Incident reports structured

__6 months → 2 days__

NLP narrative processing time

__~95 Active Users__

EXIS\.ai RAG knowledge engine

__$0 New Budget__

All tools built on existing stack

The guiding model: build freight trains, not Formula One cars\. Robust, repeatable workflows that scale across use cases and users — not one\-off demos\.

# __Operating Stack & Mindset__

TSA’s existing environment is more capable than most realize\. The constraint is not the stack — it is the mental model\.

__Layer__

__Tools__

__Workflow & Data__

Excel, SharePoint, Power Automate, PowerPoint, Teams

__AI Engine__

DHSChat — analysis, summarization, transformation, structured output

__Automation__

VBA scripting to connect workflow layers and handle data translation

__External Data__

FRED API for economic indicators

__The Three Mental Model Shifts__

- Treat AI as a process engine, not a chat toy
- Start from the business workflow — inputs, decisions, outputs — and work backward
- Encode repeatable workflows as standard patterns so they can be reused, scaled, and handed off

*“No Python\. No external infrastructure\. No budget\. Just the stack we already had and the discipline to connect it\.”*

# __Spotlight: The NLP Narrative Pipeline__

This single project most clearly illustrates the gap between what TSA’s existing tools can do and what leadership has assumed is possible\.

__The Problem__

16,000\+ free\-text narratives sat in a database — unstructured, unanalyzed, and effectively locked\. Each record contained rich operational intelligence buried in paragraph\-form text\. Manual extraction would have required an estimated 6 months of data entry work\.

__What Was Built__

A MapReduce prompt workflow using DHSChat that processed narratives in batches, extracted structured metadata from each one, and output clean JSONL — a machine\-readable format ready for analysis\. The workflow handled demographic data, operational details, intent classification, and summary generation simultaneously\.

__The Result__

__Processing time__

2 working days \(vs\. estimated 6 months manual\)

__Records processed__

16,000\+ narratives

__Output format__

Structured JSONL database, analysis\-ready

__Budget__

$0 — DHSChat only

__The Multiplier Effect__

The workflow was not hoarded\. It was documented, handed off, and taught to a colleague economist who is now independently running the same pipeline on a separate dataset\. This is the difference between a personal productivity hack and an institutional capability\.

One workflow\. Two analysts\. Months of work compressed into days\. Ready to scale further\.

# __Deployed Capabilities__

## __1\. EXIS\.ai — RAG Knowledge Engine__

__Status: Deployed and in active use | ~95 users__

Critical EXIS information was locked behind a complex interface, slowing users who needed quick answers about policies, procedures, and technical details\.

EXIS content is structured and loaded into a DHSChat Workspace so users can ask plain\-language questions and receive concise answers with accurate citations\. A key technical pain point — mis\-cited sources — was solved by redesigning how data is chunked before ingestion\.

__Active users__

~95 as of mid\-2025

__Recognition__

Cited by DHSChat development team and DHS S&T as a leading internal RAG example

__Core technique__

Custom ingestion architecture optimized for citation accuracy — solved the mis\-citation problem that plagues most RAG deployments

## __2\. Incident Report MapReduce Pipeline__

__Status: Deployed | Evidence Act support__

18,800 free\-text incident reports \(3–4 paragraphs each\) containing critical compliance data were unstructured and inaccessible for macro\-level analysis\.

A batch\-reduce workflow processed 10 reports at a time, extracting structured metadata — demographics, flight status, firearm details, intent classification, and a summary flag — for each record\. All output was re\-ingested into a structured Excel database\.

__Records processed__

18,800 incident reports

__Output__

Fully structured Excel database with granular metadata

__Technique__

Parallel\-processing batch\-reduce workflow

## __3\. Automated Economic Analysis & Modeling Pipeline__

__Status: Deployed__

Cost\-Benefit Analysis drafting is analytically intensive and document\-heavy\. Pulling foundational economic data, modeling it, and drafting the written analysis are all time\-consuming steps performed sequentially\.

An end\-to\-end automation pipeline connects directly to the FRED API to pull economic data, exports tables into JSON for the LLM, uses an engineered system prompt to draft the written CBA, and translates LLM output into JSON snippets that map to a VBA/Markdown structure — building the economic model natively inside Excel\.

__Stack__

FRED API, VBA, JSON, DHSChat, Excel

__Output__

Fully formatted Excel economic model \+ drafted CBA narrative

__Analyst impact__

Eliminates the sequential data\-pull, modeling, and drafting steps from the CBA workflow — analyst effort shifts from assembly to judgment

## __4\. Briefing Builder — Automated Presentation Engine__

__Status: Deployed__

Building standard briefings from dense source documents is repetitive and time\-consuming\. Analysts spend hours formatting slides instead of focusing on analysis and recommendations\.

A streamlined Excel UI accepts audience, purpose, length, and key emphasis inputs\. It generates an optimized prompt the user runs against source content \(40\+ page PDFs, reports, memos\)\. VBA processes the JSON output, locates the official PowerPoint template, and populates a fully branded, ~20\-slide deck with detailed presenter notes\.

__Time to produce__

Under 4 minutes for a fully formatted, TSA\-branded deck

__Output__

~20 slides with presenter notes, consistent branding

__Stack__

Excel UI, VBA, JSON, DHSChat, PowerPoint

## __5\. CFR Anti\-Competitiveness Review Engine__

__Status: Completed | Methodology reusable | Cleared legal review__

TSA needed a review of CFR sections for potential anti\-competitive language aligned with Executive Order 14036\. Traditional contractor review was time\-consuming and expensive\.

An AI\-augmented workflow parsed CFR text into reviewable format, ran section\-by\-section analysis against defined competition criteria via DHSChat, and generated structured outputs with EO\-aware summaries ready for Excel\. A human\-in\-the\-loop layer interpreted intent, market implications, and caught items AI alone would miss\. Cleared review by the Regulatory Affairs Practice Group \(RAPG\), validating the methodology for future regulatory work\.

__Sections reviewed__

378 CFR sections

__Total internal cost__

17 hours \(~$1,260 fully loaded\)

__Additional findings__

10 relevant sections identified beyond initial indications, zero false positives

__Validation__

Cleared legal review by RAPG

## __6\. Air Cargo Notice & Comment Analyzer__

__Status: Deployed | Active use__

Security Directive comment periods generate hundreds of industry responses that must be adjudicated systematically\. Manual review is slow and inconsistent\.

A MapReduce workflow processed 850 individual industry comments, automatically prioritized them \(high vs\. low\), and identified duplicates to consolidate review burden\.

__Reusable architecture surfaced: __The project produced an “Answer Key” model — distilling hundreds of comments into ~60 core questions, routing them to specific teams \(Legal, SecOps, Policy\) for pre\-adjudication, then using the finalized answers to auto\-draft legally consistent, uniform responses\. Applicable to any high\-volume comment or feedback process\.

__Comments processed__

850 industry comments

__Output__

Prioritized, deduplicated review set \+ answer\-key architecture

## __7\. RFI Sifter — Checkpoint of the Future__

__Status: Deployed on live RFIs__

RFIs generate long, inconsistent narrative responses\. Manual review makes it difficult to compare vendors and options quickly or defensibly\.

A workflow extracts key fields from each RFI response \(capabilities, technology maturity, timelines, dependencies, risk areas\) and outputs a structured comparison table\. Enables rapid filtering and prioritization based on factors leadership cares about\.

__Output__

Single structured dataset from unstructured vendor narratives

__Value__

Defensible, comparable vendor landscape for acquisition and strategy decisions

## __8\. Automated Document Reviewer__

__Status: Deployed__

Policy draft peer review is time\-consuming and inconsistently applied\. Reviewers catch different things; turnaround is slow\.

Takes drafted write\-ups, processes them against a configured review standard, and uses a VBA/JSON output layer to insert suggested edits directly into Microsoft Word as native Track Changes — allowing traditional human adjudication of every suggestion\.

__Stack__

VBA, JSON, DHSChat, Microsoft Word

__Output__

Native Word Track Changes — no new review interface required

## __9\. Rapid Survey Analyzer__

__Status: In routine use as a repeatable pattern__

Staff surveys and stakeholder feedback generate hundreds of free\-text responses\. Manual analysis is slow, subjective, and delays actionable insights\.

Starting from an Excel export, DHSChat identifies major themes and sub\-themes, quantifies frequency, surfaces representative quotes, and produces a narrative summary tailored to leadership\. Repeatable — each new survey is plug\-and\-play\.

__Time savings__

Days to weeks reduced to hours

__Output__

Quantified themes \+ narrative summary, executive\-ready

## __10\. HQ Desk Reservation Power App__

__Status: Deployed | In active use across multiple HQ floors__

The January 2025 Return\-to\-Office mandate created a severe desk shortage at HQ\. A six\-figure software contract to address it was canceled, leaving no solution\.

A Microsoft Power App built in 16 hours allows users across multiple floors to release their desks while on leave or travel, opening inventory for others to reserve dynamically\.

__Build time__

16 hours

__Cost__

$0

__Replaced__

Canceled six\-figure enterprise software contract

__Coverage__

Multiple floors of TSA HQ

## __11\. Transcription\-Based Workflow Suite__

__Status: In regular use as repeatable patterns__

Meetings, brainstorming sessions, and presentations generate spoken content that is difficult to capture, organize, and act on\. Knowledge walks out the door after every session\.

A family of patterns built on the same core idea:

- Meeting Minutes Generator — Teams transcript in, structured minutes out \(decisions, action items with owners and dates, open issues\)
- Thought Organizer — verbal SME brain dump in, coherent outline or first\-draft document out
- Whiteboard Brainstorm Converter — photo of whiteboard in, structured themes and next steps out
- Presentation Coach — slide deck \+ practice transcript in, targeted feedback on clarity, pacing, and gaps out

__Use cases__

4 repeatable patterns covering meetings, SME capture, brainstorming, and presentation prep

__Stack__

Teams, DHSChat, Excel

__Output__

Structured minutes, outlines, and coaching notes — spoken knowledge converted to actionable artifacts

# __Ready\-to\-Scale Concepts__

These capabilities are technically proven and validated\. They can be formalized and scaled quickly when leadership identifies the demand signal\.

__Capability__

__Description__

__Automated Word Memo Generator__

Converts structured inputs and source material into fully formatted memos based on official templates

__Excel Data Visualizer__

Automatically generates standardized charts and tables from selected data ranges for recurring report types

__Integrated Multi\-Step Workflows__

Chains analysis → summary → briefing → task list into one end\-to\-end flow using proven patterns above

# __Value Proposition__

Across this portfolio, the following capabilities have been consistently demonstrated:

- __Identify high\-friction processes — research, analysis, documentation, briefings — and redesign them with AI__
- __Deliver working solutions inside existing constraints: no new systems, no new budget, no new approvals__
- __Combine AI with human expertise in a way that improves both speed and quality, not one at the expense of the other__
- __Document methods so they can be reused, scaled, and handed off — turning one\-off wins into repeatable institutional capabilities__
- __Multiply impact by enabling colleagues: the goal is institutional capability, not personal productivity__

*The patterns are proven\. The stack is already in place\. TSA doesn’t need to build this from scratch — it needs to recognize what already exists and scale it\.*

