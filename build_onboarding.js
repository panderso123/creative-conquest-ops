const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,Header,Footer,
 AlignmentType,LevelFormat,HeadingLevel,BorderStyle,WidthType,ShadingType,PageNumber}=require('docx');

const NAVY="1F3A5F", GOLD="B08D57", GREY="6E6E6E", LIGHT="F2EEE6";

const styles={
  default:{document:{run:{font:"Arial",size:21,color:"222222"}}},
  paragraphStyles:[
    {id:"Heading1",name:"Heading 1",basedOn:"Normal",next:"Normal",quickFormat:true,
      run:{size:30,bold:true,color:NAVY,font:"Arial"},
      paragraph:{spacing:{before:320,after:160},outlineLevel:0,
        border:{bottom:{style:BorderStyle.SINGLE,size:6,color:GOLD,space:4}}}},
    {id:"Heading2",name:"Heading 2",basedOn:"Normal",next:"Normal",quickFormat:true,
      run:{size:24,bold:true,color:NAVY,font:"Arial"},
      paragraph:{spacing:{before:200,after:100},outlineLevel:1}},
  ]
};

const numbering={config:[
  {reference:"check",levels:[{level:0,format:LevelFormat.BULLET,text:"☐",alignment:AlignmentType.LEFT,
    style:{run:{font:"Arial",size:22},paragraph:{indent:{left:600,hanging:340}}}}]},
]};

const check=(t,bold=false)=>new Paragraph({numbering:{reference:"check",level:0},
  spacing:{after:60},
  children:[new TextRun({text:t,bold})]});
const para=(t,o={})=>new Paragraph({spacing:{after:o.after??120},children:[new TextRun({text:t,italics:o.it,color:o.color,bold:o.bold,size:o.size})]});
const h1=t=>new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun(t)]});
const h2=t=>new Paragraph({heading:HeadingLevel.HEADING_2,children:[new TextRun(t)]});

// info table
const cellB={style:BorderStyle.SINGLE,size:1,color:"D9D2C5"};
const cb={top:cellB,bottom:cellB,left:cellB,right:cellB};
function infoRow(label,val){
  return new TableRow({children:[
    new TableCell({borders:cb,width:{size:2800,type:WidthType.DXA},shading:{fill:LIGHT,type:ShadingType.CLEAR},
      margins:{top:80,bottom:80,left:140,right:140},
      children:[new Paragraph({children:[new TextRun({text:label,bold:true,color:NAVY,size:20})]})]}),
    new TableCell({borders:cb,width:{size:6560,type:WidthType.DXA},
      margins:{top:80,bottom:80,left:140,right:140},
      children:[new Paragraph({children:[new TextRun({text:val,color:GREY})]})]}),
  ]});
}
const info=new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[2800,6560],rows:[
  infoRow("Client / Practice","_______________________________"),
  infoRow("Primary contact","_______________________________"),
  infoRow("Email / phone","_______________________________"),
  infoRow("Tier","Foundation  /  Growth  /  Authority"),
  infoRow("Deposit paid","_____________   ($1,500 reserves slot, content day & launch pricing)"),
  infoRow("Start date","_____________"),
  infoRow("Content day (date)","_____________"),
  infoRow("Owner (seat)","Account & Content Lead"),
]});

const children=[
  new Paragraph({children:[new TextRun({text:"CREATIVE CONQUEST",bold:true,color:GOLD,size:18})]}),
  new Paragraph({spacing:{after:60},children:[new TextRun({text:"Client Onboarding Checklist",bold:true,color:NAVY,size:40})]}),
  para("Run this the moment a deposit clears. Target: client fully set up and first content day booked within 14 days. Owner = Account & Content Lead (the Conductor).",{it:true,color:GREY,after:200}),
  info,
  new Paragraph({spacing:{after:120},children:[]}),

  h1("1 · Pre-Kickoff (within 24 hours of deposit)"),
  check("Send welcome email + receipt and confirm the reserved slot",true),
  check("Create client folder in Google Drive (Raw / Edited / Brand / Approved)"),
  check("Create client vault entry in Bitwarden for shared logins"),
  check("Add client + deal to GoHighLevel pipeline; move stage to 'Onboarding'"),
  check("Send the intake form (links below) and brand-asset request"),
  check("Schedule the kickoff call (30 min) and the content day"),

  h1("2 · Access & Logins"),
  para("Collect via Bitwarden — never over email or text. Request admin or editor access, not the client's personal password where a team seat exists.",{it:true,color:GREY}),
  check("Instagram (Reels) — login or added as collaborator"),
  check("TikTok — login or team access"),
  check("YouTube (Shorts) — channel manager access"),
  check("LinkedIn (Authority tier only)"),
  check("Meta Business Suite / Buffer connected for scheduling"),
  check("Confirm 2FA recovery method is documented in the vault"),

  h1("3 · Brand Kit & Assets"),
  check("Logo files (PNG + vector)",true),
  check("Brand colors (hex) and fonts"),
  check("Existing photos / b-roll / past videos"),
  check("Tone & voice notes — what they will and won't say"),
  check("Treatment / service list with correct medical terminology"),
  check("Compliance no-go list (claims they cannot make)"),
  check("Build the Canva brand kit (colors, fonts, caption templates, thumbnail templates)"),

  h1("4 · Strategy & Content Plan"),
  check("Confirm target client & the 'gap we fill' (premium, well-reviewed, invisible on video)",true),
  check("Agree the content pillars (Education, Myth-Buster, Demo, Patient Result, FAQ, etc.)"),
  check("Set platform mix per tier (Foundation 1+repurpose, Growth 2-3, Authority all + LinkedIn)"),
  check("Set monthly volume (Foundation 12 / Growth 20 / Authority 30+)"),
  check("Draft first 20 video concepts in the Content Production Tracker"),
  check("Confirm posting cadence (~5/week) and best posting windows"),

  h1("5 · Content Day Logistics"),
  check("Confirm date, location, and on-camera talent (provider)",true),
  check("Send shot list & talking-point prompts in advance"),
  check("Equipment ready: phone, $30 tripod, clip mic, lighting check"),
  check("Wardrobe / setting / signage guidance sent"),
  check("Plan to film a full month (20+) in one session"),
  check("Backup: footage uploaded to Drive same day"),

  h1("6 · Systems Setup"),
  check("Stripe / PayPal subscription set to monthly tier price",true),
  check("Recurring invoice / billing date confirmed"),
  check("GoHighLevel calendar + pipeline live for this client"),
  check("ROI report template duplicated and labeled for this client"),
  check("Engagement window scheduled (daily comments & first-reply DMs)"),

  h1("7 · Kickoff Call Agenda (30 min)"),
  check("Re-state the promise: booked calendar + bought-back time, no ad spend",true),
  check("Walk through the 90-day arc and what month 1 looks like"),
  check("Confirm content day and what they need to prepare"),
  check("Set expectations: they film once/month, we run the engine"),
  check("Restate the guarantee (satisfaction-based on quality — never an income promise)"),
  check("Agree the monthly ROI report + call cadence"),

  h1("8 · Week 1 — First Proof"),
  check("First batch edited in CapCut and polished in Canva",true),
  check("QA pass against the brand standard"),
  check("Client approval on first 3-5 videos"),
  check("Scheduled in Buffer / Meta Suite"),
  check("First post live — overdeliver early"),
  check("Move GoHighLevel stage to 'Active / Delivering'"),

  new Paragraph({spacing:{before:240},border:{top:{style:BorderStyle.SINGLE,size:6,color:GOLD,space:4}},
    children:[new TextRun({text:"“Where focus goes, energy flows.”",italics:true,color:GREY})]}),
];

const doc=new Document({styles,numbering,sections:[{
  properties:{page:{size:{width:12240,height:15840},margin:{top:1440,right:1440,bottom:1440,left:1440}}},
  footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,
    children:[new TextRun({text:"Creative Conquest · Client Onboarding · Page ",size:16,color:GREY}),
              new TextRun({children:[PageNumber.CURRENT],size:16,color:GREY})]})]})},
  children}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync("Client_Onboarding_Checklist.docx",b);console.log("saved onboarding")});
