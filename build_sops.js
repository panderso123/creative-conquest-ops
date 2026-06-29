const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,Footer,
 AlignmentType,LevelFormat,HeadingLevel,BorderStyle,WidthType,ShadingType,PageNumber,PageBreak,TableOfContents}=require('docx');

const NAVY="1F3A5F",GOLD="B08D57",GREY="6E6E6E",LIGHT="F2EEE6";

const styles={
  default:{document:{run:{font:"Arial",size:21,color:"222222"}}},
  paragraphStyles:[
    {id:"Heading1",name:"Heading 1",basedOn:"Normal",next:"Normal",quickFormat:true,
      run:{size:30,bold:true,color:NAVY,font:"Arial"},
      paragraph:{spacing:{before:300,after:140},outlineLevel:0,
        border:{bottom:{style:BorderStyle.SINGLE,size:6,color:GOLD,space:4}}}},
    {id:"Heading2",name:"Heading 2",basedOn:"Normal",next:"Normal",quickFormat:true,
      run:{size:23,bold:true,color:NAVY,font:"Arial"},
      paragraph:{spacing:{before:180,after:80},outlineLevel:1}},
  ]
};
const numbering={config:[
  {reference:"steps",levels:[{level:0,format:LevelFormat.DECIMAL,text:"%1.",alignment:AlignmentType.LEFT,
    style:{paragraph:{indent:{left:620,hanging:360}}}}]},
  {reference:"steps2",levels:[{level:0,format:LevelFormat.DECIMAL,text:"%1.",alignment:AlignmentType.LEFT,
    style:{paragraph:{indent:{left:620,hanging:360}}}}]},
  {reference:"steps3",levels:[{level:0,format:LevelFormat.DECIMAL,text:"%1.",alignment:AlignmentType.LEFT,
    style:{paragraph:{indent:{left:620,hanging:360}}}}]},
  {reference:"steps4",levels:[{level:0,format:LevelFormat.DECIMAL,text:"%1.",alignment:AlignmentType.LEFT,
    style:{paragraph:{indent:{left:620,hanging:360}}}}]},
  {reference:"dot",levels:[{level:0,format:LevelFormat.BULLET,text:"•",alignment:AlignmentType.LEFT,
    style:{paragraph:{indent:{left:620,hanging:340}}}}]},
]};
// dedupe bullet refs per usage to avoid restart issues: use distinct dot refs
for(let i=1;i<=8;i++)numbering.config.push({reference:"dot"+i,levels:[{level:0,format:LevelFormat.BULLET,text:"•",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:620,hanging:340}}}}]});

const h1=t=>new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun(t)]});
const h2=t=>new Paragraph({heading:HeadingLevel.HEADING_2,children:[new TextRun(t)]});
const p=(t,o={})=>new Paragraph({spacing:{after:o.after??120},children:[new TextRun({text:t,italics:o.it,color:o.color,bold:o.bold})]});
const step=(ref,t)=>new Paragraph({numbering:{reference:ref,level:0},spacing:{after:60},children:[new TextRun(t)]});
const dot=(ref,t)=>new Paragraph({numbering:{reference:ref,level:0},spacing:{after:50},children:[new TextRun(t)]});

// meta table for each SOP
const cellB={style:BorderStyle.SINGLE,size:1,color:"D9D2C5"};
const cb={top:cellB,bottom:cellB,left:cellB,right:cellB};
function meta(owner,trigger,tools){
  function row(l,v){return new TableRow({children:[
    new TableCell({borders:cb,width:{size:2200,type:WidthType.DXA},shading:{fill:LIGHT,type:ShadingType.CLEAR},
      margins:{top:60,bottom:60,left:120,right:120},
      children:[new Paragraph({children:[new TextRun({text:l,bold:true,color:NAVY,size:19})]})]}),
    new TableCell({borders:cb,width:{size:7160,type:WidthType.DXA},
      margins:{top:60,bottom:60,left:120,right:120},
      children:[new Paragraph({children:[new TextRun({text:v,size:20,color:"222222"})]})]}),
  ]});}
  return new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[2200,7160],
    rows:[row("Owner",owner),row("Trigger / cadence",trigger),row("Tools",tools)]});
}
const spacer=new Paragraph({spacing:{after:80},children:[]});

const children=[
  new Paragraph({children:[new TextRun({text:"CREATIVE CONQUEST · BUSINESS IN A BOX",bold:true,color:GOLD,size:18})]}),
  new Paragraph({spacing:{after:60},children:[new TextRun({text:"Core Operating SOPs",bold:true,color:NAVY,size:40})]}),
  p("The delegatable engine. Each procedure is written so the VA can run it without you. If it isn't written, it can't be delegated.",{it:true,color:GREY,after:120}),
  p("You keep: filming direction, sales calls, the offer, and client relationships. Everything below moves to the VA — lowest-value work first.",{color:GREY,after:160}),
  new Paragraph({children:[new TextRun({text:"Buyback order: editing & captions → scheduling → engagement/DMs → research & calendar → prospecting drafts.",bold:true,color:NAVY})]}),
  new Paragraph({spacing:{before:200,after:80},children:[new TextRun({text:"Contents",bold:true,color:NAVY,size:24})]}),
  new TableOfContents("Contents",{hyperlink:true,headingStyleRange:"1-1"}),
  new Paragraph({children:[new PageBreak()]}),

  // SOP 1
  h1("SOP 01 · Video Editing & Captions"),
  meta("Content Producer / VA","Each batch of raw footage lands in Drive (after a content day)","CapCut (edit), Canva (brand polish, captions, thumbnails)"),
  spacer,
  h2("Definition of done"),
  p("On-brand, captioned, hook in the first 1.5 seconds, exported vertical 9:16, named correctly, and sitting in the 'Edited' folder ready for QA."),
  h2("Steps"),
  step("steps","Pull the next raw clip from Drive › Raw. Log it in the Production Tracker (set Edit Status = Editing)."),
  step("steps","Trim to a single idea. Open on the strongest hook; cut all dead air in the first 3 seconds."),
  step("steps","Target 20–45 seconds. If two ideas exist, split into two videos."),
  step("steps","Add captions (auto-caption, then proofread — especially treatment names and medical terms)."),
  step("steps","Apply the Canva brand kit: fonts, colors, lower-thirds, and a branded end card."),
  step("steps","Add light b-roll / text emphasis on the key claim; keep pacing tight."),
  step("steps","Export 1080×1920, MP4. Name: Client_Pillar_ShortTitle_v1."),
  step("steps","Move to Drive › Edited. Set Edit Status = Done and QA = Pending in the tracker."),
  h2("Quality bar (the founder's standard)"),
  dot("dot1","No claim the client can't legally or medically make."),
  dot("dot1","Captions readable on mute; no typos in clinical terms."),
  dot("dot1","Looks like the best content the brand has ever produced — that's the guarantee."),

  new Paragraph({children:[new PageBreak()]}),
  // SOP 2
  h1("SOP 02 · Scheduling & Posting"),
  meta("VA","After QA pass + client approval","Buffer or Meta Business Suite; native apps as backup"),
  spacer,
  h2("Definition of done"),
  p("~5 approved videos per week scheduled to the right platforms at peak windows, captions and hashtags set, tracker updated to Scheduled/Posted."),
  h2("Steps"),
  step("steps2","Confirm the video is QA = Pass and Client Approved = Approved in the tracker. Never schedule unapproved content."),
  step("steps2","Write the caption: hook line, value, soft CTA (e.g., 'DM us to book a consult'). Match the client's voice."),
  step("steps2","Add 3–5 relevant, non-spammy hashtags and the location tag where useful."),
  step("steps2","Schedule per the tier's platform mix (Foundation 1+repurpose, Growth 2–3, Authority all + LinkedIn)."),
  step("steps2","Post into the best windows; spread ~5/week so no day is empty and none is flooded."),
  step("steps2","Log the scheduled date, platform(s), and post link in the tracker. Set Posted? = Scheduled, then Posted when live."),
  h2("Guardrails"),
  dot("dot2","Repurpose, don't duplicate awkwardly — reframe captions per platform."),
  dot("dot2","If anything looks off-brand at this stage, send back to QA — do not post to hit a quota."),

  new Paragraph({children:[new PageBreak()]}),
  // SOP 3
  h1("SOP 03 · Engagement & DMs"),
  meta("VA (first-reply) → founder/closer for booking","Daily — once or twice","Native apps + GoHighLevel for lead capture"),
  spacer,
  h2("Definition of done"),
  p("Every comment and DM gets a fast, on-brand first reply; genuine leads are captured in GoHighLevel and routed to the closer; nothing sits unanswered overnight."),
  h2("Steps"),
  step("steps3","Once or twice daily, sweep comments and DMs across all active platforms."),
  step("steps3","Reply to comments warmly and on-brand; like and respond to questions to lift reach."),
  step("steps3","For inbound DMs, send the approved first-reply (answer the question, then offer the next step)."),
  step("steps3","Spend ~10–15 min proactively engaging the client's ideal audience (local, relevant accounts)."),
  step("steps3","Tag any genuine consult interest as a lead in GoHighLevel and notify the closer."),
  step("steps3","Log DMs/inquiries and consult interest so they flow into the monthly ROI report."),
  h2("Escalation"),
  dot("dot3","Booking conversations, pricing, and medical questions → hand to the founder/closer. The VA opens the door; the closer books."),
  dot("dot3","Complaints or anything sensitive → escalate to the Managing Director immediately, do not improvise."),

  new Paragraph({children:[new PageBreak()]}),
  // SOP 4
  h1("SOP 04 · Prospecting & Outreach"),
  meta("VA drafts → founder approves voice & sends","Daily — 20 touches/day","Prospect scorer output, GoHighLevel, email/DM"),
  spacer,
  h2("Definition of done"),
  p("20 personalized outreach touches go out per day to scored A-list prospects; every touch and reply is logged; booked calls land on the founder's calendar."),
  h2("Steps"),
  step("steps4","Work the A-list first: premium + well-reviewed (4.7★+) + invisible on short-form video (no Reels / posts rarely)."),
  step("steps4","Personalize each message — reference their actual strength (e.g., a strong blog topic) and the gap (it's not on video)."),
  step("steps4","Lead with the outcome: turn their expertise into short-form video that books consults — no ad spend, they film once a month."),
  step("steps4","Send 20 touches/day across DM and email; vary the opener, keep the core pitch consistent."),
  step("steps4","Log every touch in GoHighLevel; set follow-up reminders (a 2nd and 3rd touch lands most replies)."),
  step("steps4","On a reply, drop in the approved objection responses; the only deadline is the slot — 3 spas/month."),
  step("steps4","Book interested prospects onto the founder's calendar for the close."),
  h2("Approved CTA"),
  p("“$1,500 deposit reserves your slot, your content day, and your launch pricing. I only take 3 spas a month. Want the link?”",{it:true,color:GREY}),
  h2("Guardrails"),
  dot("dot4","The VA drafts and sends in the founder's approved voice; the founder owns the close and never delegates pricing."),
  dot("dot4","Never promise specific results or income — guarantee the quality of the work only."),

  new Paragraph({spacing:{before:240},border:{top:{style:BorderStyle.SINGLE,size:6,color:GOLD,space:4}},
    children:[new TextRun({text:"Document the SOP before you hand it over — if it isn't written, it can't be delegated.",italics:true,color:GREY})]}),
];

const doc=new Document({styles,numbering,sections:[{
  properties:{page:{size:{width:12240,height:15840},margin:{top:1440,right:1440,bottom:1440,left:1440}}},
  footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,
    children:[new TextRun({text:"Creative Conquest · Core Operating SOPs · Page ",size:16,color:GREY}),
              new TextRun({children:[PageNumber.CURRENT],size:16,color:GREY})]})]})},
  children}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync("Core_Operating_SOPs.docx",b);console.log("saved sops")});
