import {useEffect,useMemo,useRef,useState} from 'react';
import * as XLSX from 'xlsx';
const UI={bg:'#0A0F1E',card:'#111827',border:'#263448',text:'#F9FAFB',sub:'#AAB5C5',teal:'#14B8A6',red:'#EF4444',yellow:'#EAB308',green:'#22C55E',blue:'#3B82F6'};
const card={background:UI.card,border:`1px solid ${UI.border}`,padding:18,borderRadius:14};
const field={width:'100%',padding:'10px 12px',background:UI.bg,color:UI.text,border:`1px solid ${UI.border}`,borderRadius:8,boxSizing:'border-box'};
const btn={background:UI.teal,color:'#062925',fontWeight:800,border:0,borderRadius:9,padding:'10px 16px',cursor:'pointer'};
const secondary={...btn,background:'#233246',color:UI.text};
const initial={workers:[],orientations:[]};
const stamp=()=>new Date().toISOString();
const id=()=>globalThis.crypto?.randomUUID?.()||`${Date.now()}-${Math.random().toString(36).slice(2)}`;
const shuffle=arr=>{const out=[...arr];for(let i=out.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1));[out[i],out[j]]=[out[j],out[i]];}return out;};
// Source: Southwest Site Orientation & Specific Guidelines, pages 1-3.
// Content must be reviewed by the safety manager before external deployment.
export const SECTIONS=[
{title:'PPE & Work Clothing',points:['Wear appropriate work clothes: no tank tops, shorts, or ripped or loose clothing.','Southwest requires high-visibility clothing and CSA-approved safety footwear with intact soles.','Wear a hard hat that is not expired, dented, or cracked; use appropriate gloves and eye protection.','Use appropriate hearing protection for high-noise work, and task-appropriate respiratory protection where required.'],questions:[['What should you do if your hard hat is cracked?',['Replace it before continuing','Cover the crack with tape','Wear it only when supervised'],0],['Which footwear meets the site guidelines?',['CSA-approved boots with intact soles','Open-toed shoes','Boots with exposed metal toes'],0],['When are safety glasses required?',['As appropriate to the task and manufacturer instructions','Only during toolbox meetings','Only when cutting metal'],0],['What should you do when a task requires a respirator?',['Use task-appropriate protection with required fit testing and clearance','Use any available dust mask','Skip it if the job is quick'],0]]},
{title:'Hygiene, Signage & Openings',points:['Smoke only in designated smoking areas or open-air top decks as permitted.','Keep eating areas and toilets clean; do not use suite toilets or bathtubs.','Obey all posted signs, caution/danger tape, and restricted-area controls.','Floor openings greater than 4 inches must be secured and marked; trades are responsible for their openings.'],questions:[['You encounter red danger tape. What should you do?',['Stay out unless authorized','Step under it if no one is nearby','Remove it to pass'],0],['You discover an unsecured floor opening. What should you do?',['Report and secure it according to site procedure','Ignore it if it is not in your work area','Cover it with loose cardboard'],0],['Where is smoking permitted?',['Only designated permitted areas','Anywhere outside','In unfinished suites'],0]]},
{title:'Conditions & Injury Reporting',points:['Report unsafe conditions and their locations to a supervisor.','Every injury on site, however minor, must be reported to the site safety advisor as soon as possible on the same day.','Workers have the right to know, participate, refuse unsafe work, and make a complaint.'],questions:[['When should an on-site injury be reported?',['As soon as possible, the same day','Only if an ambulance is needed','At the end of the month'],0],['What should you do if you notice an unsafe condition?',['Report its location and nature to a supervisor','Assume someone else will fix it','Wait for the next orientation'],0],['Which is a worker right?',['The right to refuse unsafe work','The right to ignore posted warnings','The right to bypass barricades'],0]]},
{title:'Emergency Equipment & Response',points:['Know the current locations of first aid kits, extinguishers, air horns, eye wash, and the site trailer radio. Confirm these during site induction.','Three short air-horn blasts summon first aid; one long blast signals evacuation.','On evacuation, proceed to the designated off-site muster point.','Each trade requires a designated first-aid-trained person on site, according to the form.'],questions:[['What does one long air-horn blast signal?',['Evacuation','Lunch break','Summon first aid'],0],['How is first aid summoned with the air horn?',['Three short blasts','One long blast','Two long blasts'],0],['Where should you go during an evacuation?',['The designated off-site muster point','Your vehicle without reporting','The nearest unfinished suite'],0]]},
{title:'Fall Protection, Access & Equipment',points:['Follow the form’s fall-protection guidance and current applicable legislation and safe work procedures.','Inspect fall-protection equipment and use controls appropriate to the work.','Use approved ladders, fully open and lock step ladders, and prevent ladder movement.','Make eye contact with equipment operators and use spotters where the operator’s view is blocked.'],questions:[['Before using fall protection equipment, what should you do?',['Inspect it and confirm it is appropriate','Assume the last worker inspected it','Use it only if you feel unsafe'],0],['When an equipment operator cannot see around a load, what is required?',['An appropriate spotter','Walking directly behind the load','No additional precautions'],0],['How should a step ladder be used?',['Fully spread and locked in place','Half folded against a wall','Balanced on loose blocks'],0]]},
{title:'Site Conduct & Safe Work Practices',points:['Attend documented weekly toolbox meetings and follow site-specific safe work procedures.','Maintain walkways, clean as you progress, and use the designated waste bins.','Harassment is not tolerated; alcohol and drugs are prohibited on site under the stated site rules.','Do not touch site lights, heaters, or breaker panels without authorization.','Know the company disciplinary procedure and raise concerns with your supervisor or safety advisor.'],questions:[['Who may touch site breakers and heaters?',['Authorized personnel only','Any worker who has used them before','Anyone with gloves'],0],['What should you do with loose material in a walkway?',['Remove or report it to keep access clear','Leave it until Friday','Move it into another walkway'],0],['What is required of weekly toolbox meetings?',['They must be documented','They are only optional when convenient','Only management attends'],0],['What does the site policy say about harassment?',['It will not be tolerated','It is permitted if no one complains','It is handled only at annual review'],0]]}
];

// v21.2: Added nine rotating, site-humour knowledge checks (9 of 30 questions).
// Correct options remain clear and based on the Southwest orientation handout.
// Safety manager must review wording before external use.
SECTIONS[0].questions.push(
  ['Your hard hat has a crack. What should you do?', ['Wrap it with electrical tape and call it good', 'Remove it from service and replace it', 'Tell the foreman it has been fine like that for years', 'Have a Monster and hope for the best'], 1],
  ['Your coworker says safety glasses are unnecessary for a quick chop-saw cut. What is the proper response?', ['Use the required eye and face protection for the task', 'I have made a hundred cuts without glasses', 'Ask the new guy to hold the material while you cut'], 0]
);
SECTIONS[1].questions.push(
  ['You find an unsecured floor opening. What should happen?', ['Put a traffic cone beside it and call it done', 'Keep people clear and have it properly secured and marked', 'Say the next trade should have covered it', 'Tell everyone to watch where they are walking'], 1],
  ['Where can you smoke on site?', ['Behind the drywall stack if nobody sees you', 'Only where permitted under site smoking rules', 'Anywhere as long as you offer the foreman a dart'], 1]
);
SECTIONS[2].questions.push(
  ['A coworker gets a minor injury and says it is no big deal. What should happen?', ['Report it to the site safety advisor as soon as possible that same day', 'Offer a dart and tell them to walk it off', 'Wait until Friday at 3:30'], 0]
);
SECTIONS[3].questions.push(
  ['You hear three short blasts of the air horn. What does that mean?', ['The foreman is calling everyone back from break', 'First aid is being summoned', 'The concrete truck is here early', 'The crew is being called to the trailer'], 1]
);
SECTIONS[4].questions.push(
  ['A coworker says he has worked near unprotected edges for 25 years without fall protection. What do you do?', ['He says he has been doing it that way for twenty-five years', 'Send the new guy first', 'Stop work until proper fall protection is in place', 'Ask whether it is really necessary for a five-minute job'], 2],
  ['A guardrail was removed to move materials and has not been replaced. What should you do?', ['Put a traffic cone there and call it a guardrail', 'Keep people clear and report it so proper protection is restored', 'Have a dart and debate whose job it is'], 1]
);
SECTIONS[5].questions.push(
  ['Someone leaves loose materials blocking a walkway. What should you do?', ['It belongs to another trade, so leave it', 'Clear the obstruction or report it so the walkway remains safe', 'Push it closer to the drywallers’ materials', 'Leave it for the next trade to sort out'], 1]
);


// v21.3: Expanded topic-tagged question bank. The learner sees two DIFFERENT topics per section.
// English canonical content comes from Southwest's orientation checklist; safety manager approval required.
const EXTRA_QUESTIONS = [
 [
 ['work-clothes','Which clothing is acceptable for site work?',['Proper fitted work clothing with sleeves and no loose or ripped material','Shorts and a tank top if it is warm','The same torn shirt you have worn for every concrete pour'],0],
 ['high-vis','When should you wear required high-visibility clothing?',['As required under Southwest site policy','Only for the site photo','When the safety manager comes by'],0],
 ['footwear','The sole of your safety boot is separating. What should you do?',['Replace or repair the unsuitable footwear before work','Wrap the sole in tape; it should last another week','Tell the foreman you have been wearing them for years'],0],
 ['gloves','How do you select gloves for your task?',['Use gloves suited to the specific hazard and work','Grab any glove regardless of the task','Use the same gloves for everything; they are all gloves'],0],
 ['hearing','When should hearing protection be worn?',['When site noise and the task require it','Only when the safety manager is walking through','Never, if you have worked here long enough'],0],
 ['eye','What does the form specify for chop-saw work?',['Appropriate eye protection and a face shield','Sunglasses alone','Turn your head at the last second'],0],
 ['respirator','What is required when a task calls for respiratory protection?',['Appropriate respiratory protection with required fit testing and clearance','A scarf from the truck','Hold your breath until you finish the cut'],0],
 ['hard-hat','Your hard hat has a deep crack. What now?',['Remove it from use and replace it','Tell the foreman it still fits, so it is fine','Keep wearing it; you have had it for twenty years'],0]
 ],
 [
 ['smoking','Where does Southwest allow smoking?',['Only in permitted designated areas under the site rules','In unfinished suites if nobody is looking','Wherever you left the pack of darts'],0],
 ['toilets','Where should workers use the washroom?',['At designated site toilet facilities, not suite toilets or bathtubs','In any vacant suite','In the first room without a door'],0],
 ['eating','How should eating facilities be left?',['Clean and free of garbage and PPE on eating surfaces','Leave the mess for the other shift','Leave it; somebody else cleans the lunchroom'],0],
 ['needles','You notice a hypodermic needle on site. What do you do?',['Do not touch it; contact the site safety advisor','Pick it up with your gloves','Slide it under a pallet'],0],
 ['red-tape','Red danger tape is across a doorway. What should you do?',['Stay outside the area unless authorized','Duck under if you are late','Move it to another doorway'],0],
 ['signs','What should you do when you see a posted site safety sign?',['Follow the posted instructions','Assume it is for the new guy','Read it later if someone asks about it'],0],
 ['openings','A floor opening is uncovered. What should happen?',['Keep clear and have it secured and marked as required','Put a traffic cone beside it and call it done','Tell the other trade to keep an eye on it'],0],
 ['tape-removal','After your authorized work, what should happen to caution tape you installed?',['Remove it when the hazard is properly resolved and removal is safe','Leave it forever','Give the tape to the porta potty committee'],0]
 ],
 [
 ['injury-report','A small cut occurs while working. What should happen?',['Report it to the site safety advisor promptly that same day','Wait until it needs stitches','Have a dart and call it healed'],0],
 ['unsafe-condition','You see loose material at an edge. What do you do?',['Keep clear and report the hazard and location','Ignore it unless it falls','Tell yourself somebody else is responsible'],0],
 ['refuse','You believe a task is unsafe. What right do you have?',['The right to refuse unsafe work under applicable procedures','The right to bypass all barricades','The right to keep quiet and let the new guy try it'],0],
 ['participate','Workers are entitled to which safety involvement?',['The right to participate in workplace health and safety','Only supervisors may raise concerns','Safety concerns can be raised only on Fridays'],0],
 ['know','You are unsure about the hazards of a task. What should you do?',['Ask for the information and safe work procedure before proceeding','Just copy whatever the guy beside you is doing','See what the new guy does'],0],
 ['complaints','How should a worker raise a safety concern or grievance?',['Use the established reporting process and safety contacts','Tell the crew at lunch and assume it gets back to safety','Wait until someone else complains first'],0],
 ['site-advisor','Who should be told about an injury, no matter how minor?',['The site safety advisor as soon as possible that day','Just tell your foreman if you feel like it','Nobody if you can still finish your shift'],0],
 ['hazard-location','What information helps supervisors respond to unsafe conditions?',['Where the hazard is and what the condition involves','Only who might get blamed','Whether you think the other trade caused it'],0]
 ],
 [
 ['horn-firstaid','What do three short air-horn blasts mean?',['First aid assistance is being summoned','The foreman wants everyone at the trailer','Someone is calling the crew back to work'],0],
 ['horn-evac','What does one long air-horn blast mean?',['Evacuate as instructed','The lunch truck arrived','Time for a dart'],0],
 ['muster','Where do workers report on evacuation?',['The designated off-site muster point','Their personal vehicle without reporting','Your truck; you can leave before traffic gets bad'],0],
 ['extinguishers','Where should you locate emergency fire extinguishers?',['Know their site locations as explained during orientation','Only in the superintendent’s office','Where you think the last crew left one'],0],
 ['firstaid-kit','What should each trade arrange regarding first aid?',['Their required first aid provisions and trained site person','One box of bandages for the entire year','One guy who says he has done first aid before'],0],
 ['eyewash','What should you know about emergency eye-wash stations?',['Their current locations and how to access them','They are only for washing lunch containers','They should be wherever the last crew put them'],0],
 ['radio','What is the role of the base radio in the site trailer?',['It can be used to summon emergency assistance','It only plays jobsite music','It is only for calling the foreman about deliveries'],0],
 ['airhorn-location','Before starting work, what should you confirm about air horns?',['Their location and the site emergency signals','Whether your foreman can hear it from the trailer','Only whether you can hear it from your work area'],0]
 ],
 [
 ['fall-inspect','Before relying on fall-protection equipment, what do you do?',['Inspect it and ensure it is suitable for the work','Borrow equipment without checking','Assume the previous worker did it'],0],
 ['edge','You need to work near an unprotected edge. What comes first?',['Follow the required fall-protection controls and safe work procedure','Trust the guy who says he has done it this way for twenty-five years','Tell the foreman fall protection slows the work down'],0],
 ['guardrail','A guardrail was removed and left open. What is the right response?',['Keep clear and report it so the required protection is restored','Put a traffic cone there and call it done','Tell the next crew to watch their step'],0],
 ['stepladder','How should a step ladder be set up?',['Fully opened with spreaders locked','Half folded against a wall','On a pile of spare blocks'],0],
 ['ladder-tie','What should be done to prevent ladder movement?',['Secure the ladder as required by the safe work procedure','Ask the new guy to hold it while checking his phone','The last guy used it like this and it was fine'],0],
 ['equipment-eye','Before walking near mobile equipment, what should you do?',['Make eye contact with the operator and follow traffic controls','Assume the operator sees you','Walk behind the machine because the operator should see you'],0],
 ['spotter','The operator cannot see around a load. What is required?',['Use a spotter according to the site safe work procedure','Shout louder from behind the load','Tell the operator to just take it slow'],0],
 ['ladder-type','Which ladders are permitted under the site form?',['Type 1 and Type 2 ladders','Any ladder found in the truck','A chair if it looks sturdy'],0]
 ],
 [
 ['toolbox','What does Southwest require of weekly toolbox meetings?',['Attendance and documented meetings as directed','It counts if everybody says they were there','They can always be skipped'],0],
 ['walkway','Materials obstruct a walkway. What should happen?',['Clear them safely or report the obstruction','Move it in front of the drywallers’ pile','Leave them because the other trade brought them'],0],
 ['bins','How should site garbage be handled?',['Separate it into the designated waste bins','Put it behind the other trade’s materials','Leave it for Friday at 3:30'],0],
 ['harassment','What does the site orientation say about harassment?',['It is not tolerated on site','It is fine if someone says it was a joke','Only applies to the new guy'],0],
 ['drugs','What does the site policy say about drugs and alcohol?',['They are prohibited under site safety rules','It is fine if the foreman says you seem okay','It is fine as long as you are not operating equipment'],0],
 ['electrical','Who may operate or alter site breakers, heaters and panels?',['Only authorized people','Whoever says they know which breaker it is','Anyone wearing work gloves'],0],
 ['lights','You notice an issue with a site heater. What should you do?',['Report it and leave unauthorized repairs to designated personnel','Try a few wires until it works','Try the switch a few times and see if it starts'],0],
 ['procedure','How should workers handle tasks with a site-specific safe work procedure?',['Follow the current procedure and supervisor guidance','Do it the way it was done twenty years ago regardless','Do it the same way you have for twenty years'],0]
 ]
];
EXTRA_QUESTIONS.forEach((newQs,i)=>{newQs.forEach(([topic,prompt,options,correct])=>SECTIONS[i].questions.push([prompt,options,correct,topic]));});

// v21.7: refined jobsite-realism pass on v21.6 original, site-orientation-specific questions.
// The worker answers only TWO different topics after each of the six sections.
const FORTUNA_QUESTION_ADDITIONS = [
  [
    [
      "The new guy shows up in shorts because it is hot. What should he do?",
      [
        "Change into clothing that meets the site rules before starting.",
        "Tell the foreman shorts count as ventilation.",
        "Keep them on until someone from the office shows up."
      ],
      0,
      "work-clothes"
    ],
    [
      "Your high-vis vest is in the truck. What do you do?",
      [
        "Get the required high-visibility clothing before entering the work area.",
        "Just work where equipment operators can see you.",
        "Tell the operator you are only going to be there a minute."
      ],
      0,
      "high-vis"
    ],
    [
      "You can see the steel toe through a worn-out boot. What is the move?",
      [
        "Replace the damaged safety footwear before working.",
        "Keep wearing them until the safety manager notices.",
        "Tell everyone they were expensive, so they must still be good."
      ],
      0,
      "footwear"
    ],
    [
      "A worker says his hard hat is fine even though it is visibly damaged. What should happen?",
      [
        "Remove it from service and replace it.",
        "Trust him; he has worn that one since 2004.",
        "Turn it around so nobody sees the damage."
      ],
      0,
      "hard-hat"
    ],
    [
      "You are handling sharp material but only have unsuitable gloves. What should you do?",
      [
        "Get gloves appropriate for the material and task.",
        "Use the gloves you found in the lunchroom.",
        "Tell the new guy the gloves are optional if he is careful."
      ],
      0,
      "gloves"
    ],
    [
      "The crew is about to make a chop-saw cut. What protection is needed?",
      [
        "The appropriate eye protection and face shield.",
        "Stand behind the guy who owns the saw.",
        "Squint; the other trade has been doing it for years."
      ],
      0,
      "eye"
    ]
  ],
  [
    [
      "Someone lights a dart in an unfinished suite. What is the rule?",
      [
        "Smoking is only allowed in the permitted areas.",
        "If the drywall is not up yet, it counts as outdoors.",
        "It is fine if the foreman wants one too."
      ],
      0,
      "smoking"
    ],
    [
      "A worker says an unfinished suite toilet is closer than the site facilities. What do you tell him?",
      [
        "Use the designated site toilet facilities.",
        "Put a handwritten RESERVED sign on the suite door.",
        "Ask the new guy whether the plumbing works yet."
      ],
      0,
      "toilets"
    ],
    [
      "The lunch table is covered in garbage and used PPE. What should happen?",
      [
        "Clean the eating area and remove the waste and PPE.",
        "Leave it; someone always cleans up after the concrete crew.",
        "Stack the garbage neatly so it looks intentional."
      ],
      0,
      "eating"
    ],
    [
      "A posted sign says the area is restricted. What do you do?",
      [
        "Follow the sign and obtain authorization before entering.",
        "Follow someone who says the restricted sign is from last week.",
        "Follow the guy who says he has been going through there all week."
      ],
      0,
      "signs"
    ],
    [
      "The shortest route crosses red danger tape. What is the correct move?",
      [
        "Stay out unless you are authorized to enter.",
        "Duck under; you only need to grab one thing.",
        "Send the new guy through to see what is happening."
      ],
      0,
      "red-tape"
    ],
    [
      "Your trade has made a floor opening. What is your responsibility?",
      [
        "Make sure it is properly secured and marked under site procedures.",
        "Put your toolbox next to it so people know.",
        "Assume the next trade will notice it."
      ],
      0,
      "openings"
    ],
    [
      "The hazard is resolved but yesterday’s caution tape is still up. What should happen?",
      [
        "Have it removed when authorized and safe to do so.",
        "Leave it until someone eventually gets curious.",
        "Move the tape to an unrelated area."
      ],
      0,
      "tape-removal"
    ]
  ],
  [
    [
      "A worker says a small cut is too minor to report. What is required?",
      [
        "Report every site injury to the safety advisor as soon as possible that day.",
        "Wait until there are at least three injuries to report together.",
        "Tell the foreman only if it ruins your shirt."
      ],
      0,
      "injury-report"
    ],
    [
      "You find unsecured material beside an open edge. What should you do?",
      [
        "Keep clear and report the hazard and exact location.",
        "Kick it a little farther from where you are standing.",
        "Say nothing because the material belongs to another trade."
      ],
      0,
      "unsafe-condition"
    ],
    [
      "You have not been told the hazards of a new task. What should you do?",
      [
        "Ask for the hazards and safe work procedure before starting.",
        "Watch the first person do it and copy whatever happens.",
        "Tell the foreman you have done something similar once."
      ],
      0,
      "know"
    ],
    [
      "A foreman insists you do work you believe is unsafe. What is your next step?",
      [
        "Use the right-to-refuse process and raise the safety concern.",
        "Do it; the foreman said it will only take five minutes.",
        "Send the new guy first and see how it goes."
      ],
      0,
      "refuse"
    ],
    [
      "Someone tells you safety concerns are management’s problem. What is true?",
      [
        "Workers have the right to participate in workplace health and safety.",
        "Only the guys in white hard hats get an opinion.",
        "It becomes your problem only after something falls."
      ],
      0,
      "participate"
    ],
    [
      "You want to raise a safety concern about an ongoing practice. How do you do it?",
      [
        "Use the established reporting process and safety contacts.",
        "Tell the crew at lunch and assume someone will pass it on.",
        "Wait until the crew has stopped talking about the last complaint."
      ],
      0,
      "complaints"
    ],
    [
      "Your supervisor asks where an unsafe condition is. What should you provide?",
      [
        "The location and a clear description of the hazard.",
        "Just say somewhere on P2 and hang up.",
        "Tell them which trade you think caused it, but not where it is."
      ],
      0,
      "hazard-location"
    ],
    [
      "Who must be notified about a minor injury sustained on site?",
      [
        "The site safety advisor as soon as possible that same day.",
        "Only the guy who drives the lunch truck.",
        "The new guy when he returns from the supply run."
      ],
      0,
      "site-advisor"
    ]
  ],
  [
    [
      "You hear three short air-horn blasts. What does that signal?",
      [
        "First aid is being summoned.",
        "The foreman is looking for his tape measure again.",
        "The concrete truck has finally found the entrance."
      ],
      0,
      "horn-firstaid"
    ],
    [
      "One long air-horn blast sounds. What must workers do?",
      [
        "Follow the evacuation procedure.",
        "Finish your cut; you are almost done.",
        "Ask everyone whether the horn was meant for their trade."
      ],
      0,
      "horn-evac"
    ],
    [
      "You evacuate and see people heading straight to their cars. Where should you go?",
      [
        "Report to the designated off-site muster point.",
        "Go home; your shift must be over.",
        "Follow the guy with the cleanest truck."
      ],
      0,
      "muster"
    ],
    [
      "Before starting work, what should you confirm about first aid equipment?",
      [
        "Know where the current first aid kits are located.",
        "Assume every foreman has one in his lunch box.",
        "Ask the new guy whether he brought a bandage."
      ],
      0,
      "firstaid-kit"
    ],
    [
      "Where should you look for a fire extinguisher in an emergency?",
      [
        "At the locations identified during the site induction.",
        "In whichever trade has the most expensive tools.",
        "Behind the pile everyone has been walking around for a month."
      ],
      0,
      "extinguishers"
    ],
    [
      "A coworker asks where the emergency eyewash is. What should you know?",
      [
        "Its current site location and how to get to it.",
        "Tell him to try the water cooler.",
        "Say the electrician probably moved it."
      ],
      0,
      "eyewash"
    ],
    [
      "What should workers understand about the site trailer radio?",
      [
        "It is one of the ways to summon emergency assistance.",
        "It is only for asking where the coffee run went.",
        "It belongs to whoever gets there first."
      ],
      0,
      "radio"
    ],
    [
      "What first-aid arrangement is required of each trade under the site form?",
      [
        "A designated first-aid-trained person on site as required.",
        "The oldest person automatically becomes the first aider.",
        "The new guy gets the job because he watches medical shows."
      ],
      0,
      "firstaid-trained"
    ]
  ],
  [
    [
      "Your harness was last used by someone else. What do you do first?",
      [
        "Inspect it and confirm it is suitable before using it.",
        "Ask whether the other guy made it home yesterday.",
        "Use it; he has been in the trade for twenty years."
      ],
      0,
      "fall-inspect"
    ],
    [
      "The required fall protection is missing, but the foreman wants the work done. What do you do?",
      [
        "Stop until the proper fall protection is in place.",
        "Tell him you have been doing this for twenty years.",
        "Send the new guy first; if he can do it, you can too."
      ],
      0,
      "edge"
    ],
    [
      "A crew removes a guardrail to bring materials through. What happens next?",
      [
        "Keep the area protected and restore the required guardrail as soon as safe.",
        "Put a traffic cone where it used to be.",
        "Leave it for the trade that comes after you."
      ],
      0,
      "guardrail"
    ],
    [
      "You find a step ladder being used half-folded against a wall. What is correct?",
      [
        "Use it fully open with spreaders locked, as required.",
        "Keep using it; the guy says it has always worked.",
        "Ask someone to lean on the bottom rung."
      ],
      0,
      "stepladder"
    ],
    [
      "An access ladder shifts when someone steps on it. What is needed?",
      [
        "Secure it against movement under the safe work procedure.",
        "Wedge a loose piece of drywall under one leg.",
        "Tell people to climb it gently."
      ],
      0,
      "ladder-tie"
    ],
    [
      "A machine operator has not acknowledged you. What do you do?",
      [
        "Stay clear until you make eye contact and can move safely.",
        "Walk behind the machine; he probably heard you.",
        "Wave at the other labourer and take your chances."
      ],
      0,
      "equipment-eye"
    ],
    [
      "A machine is reversing and the operator cannot see the route. What is required?",
      [
        "Use an appropriate spotter under the site procedure.",
        "Shout louder and keep walking behind it.",
        "Ask the new guy to stand where nobody can see him."
      ],
      0,
      "spotter"
    ]
  ],
  [
    [
      "You finish your work and leave offcuts blocking the walkway. What should you do?",
      [
        "Clean up the obstruction and keep the walkway clear.",
        "Leave it; the labourers need something to do.",
        "Push it over by the drywallers’ material."
      ],
      0,
      "walkway"
    ],
    [
      "You have a pile of jobsite garbage at the end of your shift. Where does it go?",
      [
        "In the designated waste bins, as required.",
        "Behind the other trade’s material stack.",
        "In the nearest empty room until someone complains."
      ],
      0,
      "bins"
    ],
    [
      "Your crew says the toolbox meeting is optional because everyone knows the rules. What applies?",
      [
        "Attend the required documented weekly toolbox meeting.",
        "Skip it; the foreman can just tell everyone later.",
        "Send the new guy to take notes for all twenty of you."
      ],
      0,
      "toolbox"
    ],
    [
      "A worker says his way of doing a task is faster than the site procedure. What matters?",
      [
        "Follow the current site-specific safe work procedure.",
        "He has twenty years of experience, so rules no longer apply.",
        "Let him finish before anyone asks questions."
      ],
      0,
      "procedure"
    ],
    [
      "You find a breaker switched off and want power back. What do you do?",
      [
        "Leave breakers to authorized personnel.",
        "Flip switches until the lights come on.",
        "Ask the drywall crew whether anyone knows which one it is."
      ],
      0,
      "electrical"
    ],
    [
      "A site heater stops working. What is the proper response?",
      [
        "Report it and leave repairs to authorized personnel.",
        "Give it a kick; that fixed the last one.",
        "Get the new guy to try a different plug."
      ],
      0,
      "lights"
    ],
    [
      "Someone says a harassing comment is fine because it was a joke. What is the rule?",
      [
        "Harassment is not tolerated on site.",
        "It is fine as long as the foreman laughed.",
        "It only counts if somebody puts it in writing."
      ],
      0,
      "harassment"
    ]
  ]
];
FORTUNA_QUESTION_ADDITIONS.forEach((questions,i)=>SECTIONS[i].questions.push(...questions));

// Retire forced punchlines; keep the actual safe answers and site procedures.
const REPLACEMENTS = [
  [/\bsmoko\b/gi,'the break'],
  [/hope gravity takes a day off/gi,'tell the foreman you have done it for twenty years'],
  [/experience beats gravity/gi,'twenty years of experience means the rule does not apply'],
  [/Trust the guy who says he has done it this way for twenty-five years/gi,'Trust the guy who says he has done it this way for twenty-five years'],
  [/Have a Monster and hope for the best/gi,'Tell the foreman it has been fine for years'],
  [/Put a traffic cone beside it and call it done/gi,'Put a traffic cone beside it and call it done'],
  [/have a dart and debate whose job it is/gi,'argue about which trade is responsible and leave it open'],
  [/Leave it for the next trade to sort out/gi,'leave it for the next trade'],
  [/Leave them because the other trade brought them/gi,'leave it for the next trade'],
  [/site pizza party/gi,'toolbox meeting'],
  [/pizza party/gi,'toolbox meeting'],
];
SECTIONS.forEach(section=>{
  section.questions=section.questions.map(q=>[
    REPLACEMENTS.reduce((t,[rx,repl])=>t.replace(rx,repl),q[0]),
    q[1].map(option=>REPLACEMENTS.reduce((t,[rx,repl])=>t.replace(rx,repl),option)),
    q[2],q[3]
  ]);
  // Keep exactly twenty questions per section while retaining the original handout-based material.
  section.questions=section.questions.slice(0,20);
});

// Group equivalent questions by hazard/topic, including earlier v21.2 questions.
export function questionTopic(q){if(q[3])return q[3];const x=q[0].toLowerCase();if(/floor opening|unsecured opening/.test(x))return 'openings';if(/smok/.test(x))return 'smoking';if(/hard hat/.test(x))return 'hard-hat';if(/injur/.test(x))return 'injury-report';if(/air.horn|blasts/.test(x))return /three short|first aid/.test(x)?'horn-firstaid':'horn-evac';if(/guardrail/.test(x))return 'guardrail';if(/walkway/.test(x))return 'walkway';if(/fall.protection|unprotected edge/.test(x))return 'edge';if(/step ladder/.test(x))return 'stepladder';if(/eye|chop.saw/.test(x))return 'eye';if(/respirator/.test(x))return 'respirator';if(/footwear|boot/.test(x))return 'footwear';return x;}
const previousQuestionPicks=new Map();
export function pickDistinctQuestions(index,count=2){
  const bank=SECTIONS[index].questions;
  const previous=previousQuestionPicks.get(index)||[];
  const unseen=shuffle(bank.filter(q=>!previous.includes(q[0])));
  const candidates=[...unseen,...shuffle(bank.filter(q=>previous.includes(q[0])))];
  const selected=[],seenTopics=new Set();
  for(const q of candidates){
    const topic=questionTopic(q);
    if(seenTopics.has(topic))continue;
    selected.push(q);seenTopics.add(topic);
    if(selected.length===count)break;
  }
  previousQuestionPicks.set(index,selected.map(q=>q[0]));
  return selected;
}
const QUESTIONS_PER_SECTION=2;
function makeAttempt(workerId){return {id:id(),worker_id:workerId,version:'Southwest Orientation v21.7 test',status:'in_progress',approval_status:'pending',started_at:stamp(),completed_at:null,sections:[],signature:'',acknowledged:false};}
function daysTo(date){if(!date)return null;const due=new Date(date+'T12:00:00'),now=new Date();now.setHours(12,0,0,0);return Math.round((due-now)/86400000);}
const cleanName=(s)=>String(s||'').trim().toLowerCase().replace(/\s+/g,' ');
function formItem(label,value,onChange,opts){return <label style={{fontSize:12,color:UI.sub,display:'grid',gap:5}}>{label}{opts?<select value={value||''} onChange={e=>onChange(e.target.value)} style={field}><option value="">Select</option>{opts.map(o=><option key={o} value={o}>{o}</option>)}</select>:<input value={value||''} onChange={e=>onChange(e.target.value)} style={field}/>}</label>}
export default function SafetyModule({onBack,onLegacy}){
const [reviewMode,setReviewMode]=useState('all');
const [showAccessKey,setShowAccessKey]=useState(false);
const [workerLink,setWorkerLink]=useState(''),[key,setKey]=useState(''),[role,setRole]=useState(''),[login,setLogin]=useState(''),[records,setRecords]=useState(initial),[rev,setRev]=useState(0),[error,setError]=useState(''),[busy,setBusy]=useState(false),[tab,setTab]=useState('dashboard'),[selected,setSelected]=useState(''),[draft,setDraft]=useState(null),[run,setRun]=useState(null),[search,setSearch]=useState(''),[filter,setFilter]=useState('all'),[directoryTrade,setDirectoryTrade]=useState(''),[directoryCompany,setDirectoryCompany]=useState('');
const activeTime=useRef({entered:0,visible:0,hidden:false});
const request=async(method,token,body)=>{const r=await fetch('/api/safety-records',{method,headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},body:body?JSON.stringify(body):undefined,cache:'no-store'});const d=await r.json();if(!r.ok)throw Error(d.error||'Request failed');return d;};
async function signIn(){setBusy(true);setError('');try{const data=await request('GET',login);setKey(login);setRole(data.role);setRecords(data.payload||initial);setRev(data.revision||0);setLogin('');}catch(e){setError(e.message)}finally{setBusy(false)}}
async function reload(){if(!key)return;setBusy(true);try{const d=await request('GET',key);setRecords(d.payload||initial);setRev(d.revision||0);setRole(d.role);setError('');}catch(e){setError(e.message)}finally{setBusy(false)}}
async function save(next){setBusy(true);setError('');try{const d=await request('POST',key,{action:'save',payload:next,revision:rev});setRecords(d.payload);setRev(d.revision);return true;}catch(e){setError(e.message);return false;}finally{setBusy(false)}}
const workers=records.workers||[],orientations=records.orientations||[];
const latestFor=(workerId)=>orientations.filter(o=>o.worker_id===workerId).sort((a,b)=>String(b.started_at).localeCompare(String(a.started_at)))[0];
const expiring=workers.flatMap(w=>(w.tickets||[]).map(t=>({...t,worker:w,remaining:daysTo(t.expiry_date)}))).filter(t=>t.remaining!==null&&t.remaining<=60);
const pending=orientations.filter(o=>o.status==='completed'&&o.approval_status==='pending');
const pendingTickets=workers.flatMap(worker=>(worker.tickets||[]).filter(ticket=>ticket.status!=='verified').map(ticket=>({worker,ticket}))).sort((a,b)=>String(b.ticket.added_at||'').localeCompare(String(a.ticket.added_at||'')));
const selectedWorker=workers.find(w=>w.id===selected);
function editWorker(w){setDraft(w?JSON.parse(JSON.stringify(w)):{id:id(),name:'',company:'',trade:'',occupation:'',supervisor:'',tickets:[],created_at:stamp(),historical_orientation:null});setTab('edit')}
async function submitWorker(){if(!draft?.name?.trim()||!draft?.company?.trim()||!draft?.trade?.trim()){setError('Worker name, company and trade are required.');return;}const duplicate=workers.some(w=>w.id!==draft.id&&cleanName(w.name)===cleanName(draft.name)&&cleanName(w.company)===cleanName(draft.company));if(duplicate){setError('This worker already exists at the same company. Open their profile instead.');return;}const next={...records,workers:[...workers.filter(w=>w.id!==draft.id),{...draft,name:draft.name.trim()}]};if(await save(next)){setSelected(draft.id);setTab('workers');setDraft(null)}}
async function addTicket(w,values,file){
 setError('');
 if(!file){setError('Select the original certificate photo or PDF.');return false;}
 if(file.size>10*1024*1024){setError('Select a file under 10 MB.');return false;}
 const mime=file.type||({'pdf':'application/pdf','jpg':'image/jpeg','jpeg':'image/jpeg','png':'image/png','webp':'image/webp'}[file.name.split('.').pop().toLowerCase()]||'');
 try{
  setBusy(true);
  const signed=await request('POST',key,{action:'certificate_upload_url',workerId:w.id,mime,size:file.size});
  const result=await fetch(signed.uploadUrl,{method:'PUT',headers:{'Content-Type':mime,'x-upsert':'false'},body:file});
  if(!result.ok)throw Error('Could not upload the original certificate. Please try again.');
  const attachment={name:file.name,type:mime,size:file.size,storage_path:signed.storagePath,bucket:'fortuna-worker-certificates'};
  // Use the most recent record to avoid overwriting a worker's concurrently submitted certificate.
  const fresh=await request('GET',key);
  const updatedWorker=(fresh.payload.workers||[]).find(x=>x.id===w.id);
  if(!updatedWorker)throw Error('Worker no longer exists. Refresh the roster.');
  const nextTicket={id:id(),type:values.type?.trim()||'Unidentified certificate — review required',issued_date:values.issued_date||'',expiry_date:values.expiry_date||'',holder_name:values.holder_name?.trim()||'',notes:values.notes||'',status:'needs_review',attachment,added_at:stamp(),source:'staff_upload'};
  const next={...fresh.payload,workers:fresh.payload.workers.map(x=>x.id===w.id?{...x,tickets:[...(x.tickets||[]),nextTicket]}:x)};
  const saved=await request('POST',key,{action:'save',payload:next,revision:fresh.revision});
  setRecords(saved.payload);setRev(saved.revision);setSelected(w.id);return true;
 }catch(e){setError(e.message||'Certificate upload failed');return false;}
 finally{setBusy(false)}
}
async function deleteWorker(w){
 if(role!=='manager')return setError('Only the safety manager can delete workers.');
 const typed=window.prompt(`PERMANENTLY delete ${w.name} (${w.company})?\n\nThis deletes their profile, certificates, and all orientation attempts.\nType the worker's full name to continue:`);
 if(typed===null)return;
 if(cleanName(typed)!==cleanName(w.name)){setError('Name did not match. Nothing was deleted.');return;}
 if(!window.confirm(`Final confirmation: permanently delete ${w.name} from ${w.company}? This cannot be undone.`))return;
 setBusy(true);setError('');
 try{const d=await request('POST',key,{action:'worker_delete',workerId:w.id,confirmName:typed,revision:rev});setRecords(d.payload);setRev(d.revision);setSelected('');setDraft(null);setWorkerLink(d.warning||`Deleted ${w.name} and associated orientation and certificate records.`);}
 catch(e){setError(e.message||'Could not delete worker');}
 finally{setBusy(false)}
}
async function reviewTicket(w,ticketId,fields){if(role!=='manager')return setError('Safety manager credentials are required to verify certificates.');setBusy(true);setError('');try{const d=await request('POST',key,{action:'ticket_review',workerId:w.id,ticketId,fields,revision:rev});setRecords(d.payload);setRev(d.revision)}catch(e){setError(e.message)}finally{setBusy(false)}}
async function viewOriginalCertificate(workerId,ticketId){
 setError('');
 try{const r=await request('POST',key,{action:'certificate_view',workerId,ticketId});window.open(r.url,'_blank','noopener,noreferrer');}
 catch(e){setError(e.message||'Cannot open original certificate');}
 }
 async function historical(w,date,reviewer,attachment){if(!date||!reviewer)return setError('Original date and verifier are required.');const updated={...w,historical_orientation:{date,verified_by:reviewer,record:attachment||null,entered_at:stamp()}};await save({...records,workers:workers.map(x=>x.id===w.id?updated:x)})}
async function startOrientation(w){const attempt=makeAttempt(w.id);if(await save({...records,orientations:[...orientations,attempt]})){setRun({attempt,section:0,started:Date.now(),questions:pickQuestions(0),selected:-1,tries:0,log:[],ack:false,signature:'',declaration:false});activeTime.current={entered:Date.now(),visible:0,hidden:document.hidden};setTab('orientation')}}
function pickQuestions(i){return pickDistinctQuestions(i,QUESTIONS_PER_SECTION).map(q=>({prompt:q[0],options:shuffle(q[1].map((text,index)=>({text,correct:index===q[2]}))),answered:false,attempts:0}));}
function elapsed(){const a=activeTime.current;return a.visible+(!a.hidden?Date.now()-a.entered:0)}
useEffect(()=>{function vis(){const a=activeTime.current;if(document.hidden&&!a.hidden){a.visible+=Date.now()-a.entered;a.hidden=true;}else if(!document.hidden&&a.hidden){a.entered=Date.now();a.hidden=false;}}document.addEventListener('visibilitychange',vis);return()=>document.removeEventListener('visibilitychange',vis)},[]);
async function nextSection(){if(!run||!run.ack||run.questions.some(q=>!q.answered))return;const event={title:SECTIONS[run.section].title,elapsed_seconds:Math.round(elapsed()/1000),questions:run.questions.map(q=>({prompt:q.prompt,attempts:q.attempts,correct:true})),acknowledged_at:stamp()};const attempt={...run.attempt,sections:[...run.attempt.sections,event]};if(run.section===SECTIONS.length-1){setRun({...run,attempt,section:SECTIONS.length,questions:[],ack:false});return;}const nextIndex=run.section+1;setRun({...run,attempt,section:nextIndex,questions:pickQuestions(nextIndex),selected:-1,tries:0,log:[],ack:false});activeTime.current={entered:Date.now(),visible:0,hidden:document.hidden}}
async function completeOrientation(){if(!run?.signature?.trim()||!run.declaration)return setError('Read and sign the declaration to finish.');const completed={...run.attempt,status:'completed',completed_at:stamp(),signature:run.signature.trim(),acknowledged:true,approval_status:'pending'};if(await save({...records,orientations:orientations.map(o=>o.id===completed.id?completed:o)})){setRun(null);setSelected(completed.worker_id);setTab('workers')}}
async function getWorkerLink(w){setBusy(true);setError('');try{const d=await request('POST',key,{action:'worker_link',workerId:w.id,revision:rev});const url=`https://fortuna-worker-orientation.vercel.app/?session=${encodeURIComponent(d.token)}`;await navigator.clipboard.writeText(url);setWorkerLink(`Secure link for ${w.name} copied — expires in 12 hours`);}catch(e){setError(e.message)}finally{setBusy(false)}}
async function review(o,decision){if(role!=='manager')return setError('Sign in using safety-manager credentials to approve orientations.');setBusy(true);try{const d=await request('POST',key,{action:'review',workerId:o.id,decision,revision:rev});setRecords(d.payload);setRev(d.revision);setError('')}catch(e){setError(e.message)}finally{setBusy(false)}}
function exportSheet(){
 const wb=XLSX.utils.book_new();
 const add=(name,rows)=>{const sheet=XLSX.utils.json_to_sheet(rows.length?rows:[{Information:'No records available'}]);XLSX.utils.book_append_sheet(wb,sheet,name);};
 add('Worker Roster',workers.map(w=>{const o=latestFor(w.id);return {Trade:w.trade||'',Company:w.company||'',Worker:w.name||'',Occupation:w.occupation||'',Supervisor:w.supervisor||'',Orientation:o?.approval_status==='approved'?'Verified':o?.status==='completed'?'Awaiting manager':w.historical_orientation?'Previously oriented':'Not oriented',OrientationDate:o?.completed_at||w.historical_orientation?.date||'',Certificates:(w.tickets||[]).length,ExpiringIn30Days:(w.tickets||[]).filter(t=>{const d=daysTo(t.expiry_date);return d!==null&&d>=0&&d<=30}).length}}));
 add('Certificates',workers.flatMap(w=>(w.tickets||[]).map(t=>({Trade:w.trade||'',Company:w.company||'',Worker:w.name||'',CertificateType:t.type||'',NameOnCertificate:t.holder_name||t.name||'',IssueDate:t.issue_date||'',ExpiryDate:t.expiry_date||'',Status:t.status||'',DateSubmitted:t.added_at||'',DaysUntilExpiry:daysTo(t.expiry_date)??''}))));
 add('Orientations',orientations.map(o=>{const w=workers.find(w=>w.id===o.worker_id);return {Worker:w?.name||'',Company:w?.company||'',Trade:w?.trade||'',Status:o.status||'',ApprovalStatus:o.approval_status||'',Started:o.started_at||'',Completed:o.completed_at||'',Signature:o.signature||'',Language:o.language||'',IncorrectTries:o.incorrect_attempts??'',SectionsCompleted:Array.isArray(o.sections)?o.sections.length:''}}));
 add('Section History',orientations.flatMap(o=>{const w=workers.find(w=>w.id===o.worker_id);return (o.sections||[]).map((section,i)=>({Worker:w?.name||'',Company:w?.company||'',Completed:o.completed_at||'',Section:section.title||section.name||('Section '+(i+1)),DurationSeconds:section.seconds??section.duration_seconds??'',IncorrectAttempts:section.incorrect_attempts??'',Questions:Array.isArray(section.questions)?section.questions.map(q=>q.prompt||q.question||'').join(' | '):''}))}));
 XLSX.writeFile(wb,'fortuna-worker-safety-report.xlsx');
}
if(!role)return <div style={{minHeight:'100vh',background:UI.bg,color:UI.text,padding:'8vh 24px',fontFamily:'Segoe UI,sans-serif'}}><div style={{...card,maxWidth:470,margin:'auto'}}><button onClick={onBack} style={secondary}>← Back</button><h2>Southwest Safety & Orientations</h2><p style={{color:UI.sub}}>Internal staff sign-in. Public worker access is available only in this test environment.</p><form onSubmit={e=>{e.preventDefault();signIn()}}><label>Staff or manager access key<div style={{display:'flex',alignItems:'center',gap:8,marginTop:8}}><input type={showAccessKey?'text':'password'} autoComplete="off" value={login} onChange={e=>setLogin(e.target.value)} style={{...field,flex:1,minWidth:0}}/><button type="button" onClick={()=>setShowAccessKey(v=>!v)} aria-label={showAccessKey?'Hide access key':'Show access key'} aria-pressed={showAccessKey} title={showAccessKey?'Hide password':'Show password'} style={{...secondary,padding:'10px 12px',whiteSpace:'nowrap'}}><span style={{display:'inline-flex',alignItems:'center',gap:6}}><svg aria-hidden="true" width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M2 12s3.6-7 10-7 10 7 10 7-3.6 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/>{!showAccessKey&&<path d="M3 3l18 18"/>}</svg>{showAccessKey?'Hide':'Show'}</span></button></div></label><button disabled={busy||!login} style={{...btn,marginTop:15}}>Sign in</button></form>{error&&<p style={{color:UI.red}}>{error}</p>}</div></div>;
const openDashboardSection=(destination,mode)=>{setError('');if(destination==='reviews')setReviewMode(mode||'all');if(destination==='expiry')setFilter(mode||'all');setTab(destination);};
return <div style={{minHeight:'100vh',background:UI.bg,color:UI.text,paddingBottom:60,fontFamily:'Segoe UI,sans-serif'}}>
<div style={{padding:'20px 26px',borderBottom:`1px solid ${UI.border}`,display:'flex',justifyContent:'space-between',flexWrap:'wrap',gap:10}}><div><button style={secondary} onClick={onBack}>← Modules</button><h2 style={{margin:'12px 0 4px'}}>Worker Safety & Orientations <small style={{color:UI.teal}}>v21.6</small></h2><div style={{color:UI.sub,fontSize:12}}>Fortuna · {role==='manager'?'Safety manager':'Southwest staff'} · Internal testing</div></div><div style={{display:'flex',gap:8,alignItems:'center',flexWrap:'wrap'}}><button style={secondary} onClick={reload}>↻ Refresh</button><button style={secondary} onClick={()=>{setRole('');setKey('')}}>Sign out</button></div></div>

{workerLink&&<div style={{margin:'10px 24px',color:UI.green}}>{workerLink}</div>}{error&&<div role="alert" style={{margin:'15px 24px',padding:12,background:'#461f26',color:'#fecaca',borderRadius:8}}>{error}</div>}{busy&&<div style={{margin:'12px 24px',color:UI.sub}}>Saving / loading…</div>}
<main style={{padding:'24px',maxWidth:1200,margin:'auto'}}>
{tab!=='dashboard'&&<button style={{...secondary,marginBottom:16}} onClick={()=>{setTab('dashboard');setError('');setSelected('');setDraft(null)}}>← Back to Overview</button>}
{tab==='dashboard'&&<><div style={{display:'flex',gap:12,flexWrap:'wrap',marginBottom:22}}>{[['Workers',workers.length,'workers'],['Orientations verified',orientations.filter(o=>o.approval_status==='approved').length,'reviews','approved'],['Orientations pending',pending.length,'reviews','pending'],['Certificates pending',pendingTickets.length,'certificates'],['Tickets expiring ≤30d',expiring.filter(t=>t.remaining<=30).length,'expiry','30']].map(([label,value,destination,mode])=><button key={label} type="button" onClick={()=>openDashboardSection(destination,mode)} aria-label={`Open ${label}`} style={{...card,minWidth:170,flex:1,textAlign:'left',color:UI.text,cursor:'pointer',fontFamily:'inherit'}}><div style={{color:UI.sub,fontSize:12}}>{label} <span style={{float:'right'}}>↗</span></div><div style={{fontSize:29,color:UI.teal,fontWeight:800}}>{value}</div></button>)}</div><div style={card}><h3>Orientation and Certificate Management</h3><p style={{color:UI.sub}}>Start internal orientations, register existing workers, save historical records, review individual tickets and export your roster. Worker uploads appear in the certificate review queue until a safety manager verifies them.</p><div style={{display:'flex',gap:10,flexWrap:'wrap'}}>{role==='manager'&&<button style={btn} onClick={()=>editWorker()}>+ Add Worker</button>}<button style={secondary} onClick={()=>setTab('reviews')}>Review Orientations ({pending.length})</button><button style={secondary} onClick={()=>setTab('certificates')}>Review Certificates ({pendingTickets.length})</button><button style={secondary} onClick={exportSheet}>Export Excel</button><a href='https://fortuna-worker-orientation.vercel.app/' target='_blank' rel='noreferrer' style={{...secondary,textDecoration:'none'}}>Open Public Worker Portal ↗</a><details style={{color:UI.sub,marginTop:12,width:'100%'}}><summary style={{cursor:'pointer'}}>Show Worker Orientation QR Code</summary><p style={{fontSize:12}}>Scan to open public worker registration. Share this general QR code, never a private returning-worker link.</p><img width='220' height='220' alt='QR code for the public Fortuna worker orientation portal' style={{background:'#fff',padding:8,maxWidth:'100%',height:'auto'}} src='data:image/png;base64,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'/><p style={{fontSize:12}}><a href='https://fortuna-worker-orientation.vercel.app/' target='_blank' rel='noreferrer' style={{color:UI.teal}}>fortuna-worker-orientation.vercel.app ↗</a></p></details></div></div>{pendingTickets.length>0&&<div style={{...card,marginTop:14}}><h3>Certificates awaiting Southwest review ({pendingTickets.length})</h3>{pendingTickets.slice(0,5).map(({worker,ticket})=><div key={`${worker.id}-${ticket.id}`} style={{display:'flex',justifyContent:'space-between',gap:12,alignItems:'center',flexWrap:'wrap',padding:'10px 0',borderBottom:`1px solid ${UI.border}`}}><span><strong>{worker.name}</strong> · {worker.company} · {ticket.type||'Unidentified certificate'}</span><button style={secondary} onClick={()=>setTab('certificates')}>Review certificate →</button></div>)}</div>}</>}
{tab==='workers'&&<>
<div style={{display:'flex',gap:9,marginBottom:18,flexWrap:'wrap'}}>
<input aria-label="Find worker, company or trade" placeholder="Search any worker, company or trade" value={search} onChange={e=>{setSearch(e.target.value);setSelected('')}} style={{...field,maxWidth:340}}/>
{role==='manager'&&<button style={btn} onClick={()=>editWorker()}>+ Add Worker</button>}<button style={secondary} onClick={exportSheet}>Export</button>
</div>
{(()=>{
const normalize=v=>String(v||'').trim();
const sorted=(values)=>[...new Set(values)].sort((a,b)=>a.localeCompare(b));
const shown=workers.filter(w=>[w.name,w.trade,w.company,w.occupation].some(v=>normalize(v).toLowerCase().includes(search.trim().toLowerCase())));
const tradeNames=sorted(shown.map(w=>normalize(w.trade)||'Unassigned'));
const current=selectedWorker;
const currentTrade=directoryTrade;
const currentCompany=directoryCompany;
const crumb=<div style={{display:'flex',alignItems:'center',gap:8,flexWrap:'wrap',marginBottom:14}}>
{(current||currentTrade)&&<button style={secondary} onClick={()=>{if(current)setSelected('');else if(currentCompany)setDirectoryCompany('');else setDirectoryTrade('')}}>← Back</button>}
<span style={{color:UI.sub,fontSize:13}}>Trades{currentTrade?' / '+currentTrade:''}{currentCompany?' / '+currentCompany:''}{current?' / '+current.name:''}</span>
</div>;
const tile=(name,count,onClick,key)=> <button key={key||name} onClick={onClick} style={{...card,width:'100%',display:'flex',justifyContent:'space-between',alignItems:'center',gap:16,textAlign:'left',cursor:'pointer',marginBottom:10,color:UI.text}}><span><strong style={{color:UI.teal,fontSize:17}}>{name}</strong><span style={{display:'block',color:UI.sub,fontSize:13,marginTop:5}}>{count} worker{count===1?'':'s'}</span></span><span style={{fontSize:20,color:UI.sub}}>›</span></button>;
const workerTile=w=>{const o=latestFor(w.id);const status=o?.approval_status==='approved'?'Verified':o?.status==='completed'?'Pending manager':w.historical_orientation?'Historical':'Not oriented';return <button key={w.id} onClick={()=>{setDirectoryTrade(normalize(w.trade)||'Unassigned');setDirectoryCompany(normalize(w.company)||'Unassigned');setSelected(w.id);setSearch('')}} style={{...secondary,width:'100%',display:'flex',justifyContent:'space-between',alignItems:'center',textAlign:'left',padding:14,marginBottom:9,gap:8}}><strong>{w.name}</strong><span style={{color:status==='Verified'?UI.green:UI.yellow}}>{status} ›</span></button>};
return <>{crumb}
{current ? <WorkerDetails key={current.id} worker={current} attempts={orientations.filter(o=>o.worker_id===current.id)} onEdit={role==='manager'?()=>editWorker(current):null} onOrient={role==='manager'?()=>startOrientation(current):null} onTicket={role==='manager'?(v,f)=>addTicket(current,v,f):null} onVerify={role==='manager'?()=>setTab('certificates'):null} onHistory={role==='manager'?(d,r,file)=>historical(current,d,r,file):null} onWorkerLink={role==='manager'?()=>getWorkerLink(current):null} onDelete={role==='manager'?()=>deleteWorker(current):null} busy={busy} onClose={()=>setSelected('')}/> :
search.trim() ? <div style={card}><h3 style={{marginTop:0}}>Search results ({shown.length})</h3>{shown.sort((a,b)=>a.name.localeCompare(b.name)).map(w=><div key={w.id}><div style={{color:UI.sub,fontSize:12,margin:'12px 0 4px'}}>{w.trade} / {w.company}</div>{workerTile(w)}</div>)}{shown.length===0&&<p style={{color:UI.sub}}>No matches found.</p>}</div> :
!currentTrade ? <><h3>Trades</h3>{tradeNames.map(t=>tile(t,shown.filter(w=>(normalize(w.trade)||'Unassigned')===t).length,()=>{setDirectoryTrade(t);setDirectoryCompany('');setSelected('')},t))}{workers.length===0&&<p style={{color:UI.sub}}>No workers yet. Add an existing worker to start.</p>}</> :
!currentCompany ? <><h3>{currentTrade} — Companies</h3>{sorted(workers.filter(w=>(normalize(w.trade)||'Unassigned')===currentTrade).map(w=>normalize(w.company)||'Unassigned')).map(c=>tile(c,workers.filter(w=>(normalize(w.trade)||'Unassigned')===currentTrade&&(normalize(w.company)||'Unassigned')===c).length,()=>{setDirectoryCompany(c);setSelected('')},c))}</> :
<div style={card}><h3 style={{marginTop:0}}>{currentCompany} — Workers</h3>{workers.filter(w=>(normalize(w.trade)||'Unassigned')===currentTrade&&(normalize(w.company)||'Unassigned')===currentCompany).sort((a,b)=>a.name.localeCompare(b.name)).map(workerTile)}</div>}
</>;
})()}
</>}
{tab==='edit'&&draft&&role==='manager'&&<div style={{...card,maxWidth:760}}><h3>{workers.some(w=>w.id===draft.id)?'Edit':'Create'} Worker</h3><div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(200px,1fr))',gap:13}}>{[['Name','name'],['Company','company'],['Trade','trade'],['Occupation','occupation'],['Supervisor','supervisor']].map(([label,k])=>formItem(label,draft[k],v=>setDraft({...draft,[k]:v})))}</div><div style={{display:'flex',gap:10,marginTop:18}}><button style={btn} onClick={submitWorker} disabled={busy}>Save Worker</button><button style={secondary} onClick={()=>setTab('workers')}>Cancel</button></div></div>}
{tab==='reviews'&&<><h3>{reviewMode==='approved'?'Verified Orientations':reviewMode==='pending'?'Orientations Pending Approval':'Safety Manager Verification'}</h3><p style={{color:UI.sub}}>Only an authorized safety-manager session can approve or return a completed orientation.</p>{orientations.filter(o=>o.status==='completed'&&(reviewMode==='all'||o.approval_status===reviewMode)).slice().reverse().map(o=>{const w=workers.find(w=>w.id===o.worker_id);return <div style={{...card,marginBottom:12}} key={o.id}><div style={{display:'flex',justifyContent:'space-between',gap:12,flexWrap:'wrap'}}><strong>{w?.name||'Worker removed'} · {w?.company||''}</strong><span style={{color:o.approval_status==='approved'?UI.green:UI.yellow}}>{o.approval_status}</span></div><p style={{color:UI.sub}}>Completed {new Date(o.completed_at).toLocaleString()} · Signed: {o.signature} · Language: {o.language||'en'} · {o.sections.reduce((a,s)=>a+s.elapsed_seconds,0)} seconds active · {o.sections.reduce((a,s)=>a+s.questions.reduce((b,q)=>b+q.attempts-1,0),0)} incorrect tries</p><details><summary>View section history and question attempts</summary>{o.sections.map((s,i)=><div key={i} style={{padding:10,background:UI.bg,margin:6,borderRadius:6}}><b>{s.title}</b> — {s.elapsed_seconds}s {s.questions.map((q,j)=><p key={j} style={{fontSize:12,color:UI.sub}}>{q.prompt}: {q.attempts} attempt(s)</p>)}</div>)}</details>{role==='manager'&&o.approval_status!=='approved'&&<div style={{display:'flex',gap:9,marginTop:14}}><button disabled={busy} style={btn} onClick={()=>review(o,'approve')}>Approve</button><button disabled={role!=='manager'||busy} style={secondary} onClick={()=>review(o,'return')}>Return for Review</button></div>}</div>})}{orientations.filter(o=>o.status==='completed'&&(reviewMode==='all'||o.approval_status===reviewMode)).length===0&&<p style={{color:UI.sub}}>No orientations in this category.</p>}</>}
{tab==='certificates'&&<><h3>Certificates Pending Review ({pendingTickets.length})</h3><p style={{color:UI.sub}}>Check the original file, correct any missing details, and verify. Only safety managers can approve certificates.</p>{pendingTickets.length===0&&<div style={card}>No certificates awaiting verification.</div>}{pendingTickets.map(({worker,ticket})=><CertificateReview onViewOriginal={viewOriginalCertificate} key={`${worker.id}-${ticket.id}`} worker={worker} ticket={ticket} busy={busy} manager={role==='manager'} onVerify={fields=>reviewTicket(worker,ticket.id,fields)}/>)}</>}
{tab==='expiry'&&<><h3>Certificate Expiry Register</h3><select style={{...field,maxWidth:230,marginBottom:12}} value={filter} onChange={e=>setFilter(e.target.value)}><option value="all">All expiring within 60 days</option><option value="expired">Expired</option><option value="30">Within 30 days</option></select>{expiring.filter(t=>filter==='expired'?t.remaining<0:filter==='30'?t.remaining<=30:true).sort((a,b)=>a.remaining-b.remaining).map(t=><div key={t.id} style={{...card,marginBottom:9,display:'flex',justifyContent:'space-between',gap:8,flexWrap:'wrap'}}><span>{t.worker.trade} / {t.worker.name} / {t.type}</span><span style={{color:t.remaining<0?UI.red:UI.yellow}}>{t.remaining<0?`Expired ${-t.remaining} days ago`:`${t.remaining} days remaining`} · {t.expiry_date}</span></div>)}</>}
{tab==='orientation'&&run&&role==='manager'&&<div style={{...card,maxWidth:680,margin:'auto'}}><div style={{color:UI.teal,fontWeight:700}}>SOUTHWEST SITE ORIENTATION · {Math.min(run.section+1,SECTIONS.length)} / {SECTIONS.length}</div>{run.section<SECTIONS.length?<><h2>{SECTIONS[run.section].title}</h2>{SECTIONS[run.section].points.map((point,i)=><p key={i} style={{background:UI.bg,padding:12,borderRadius:7,lineHeight:1.6}}>{point}</p>)}<label style={{display:'flex',gap:8,margin:'20px 0'}}><input type="checkbox" checked={run.ack} onChange={e=>setRun({...run,ack:e.target.checked})}/> I have read and understand this section.</label><h3>Knowledge check</h3>{run.questions.map((q,i)=><div key={q.prompt} style={{...card,marginBottom:12}}><p><b>{i+1}. {q.prompt}</b></p>{q.options.map((answer,j)=><button key={j} disabled={q.answered} onClick={()=>setRun(r=>({...r,questions:r.questions.map((question,n)=>n===i?{...question,attempts:question.attempts+1,answered:answer.correct}:question)}))} style={{...secondary,display:'block',width:'100%',textAlign:'left',marginBottom:6,opacity:q.answered?.7:1}}>{answer.text}</button>)}<span style={{fontSize:12,color:q.answered?UI.green:q.attempts?UI.yellow:UI.sub}}>{q.answered?'Correct ✓':q.attempts?'Not quite. Review the section and try again.':'Choose the best answer.'}</span></div>)}<button style={btn} disabled={!run.ack||run.questions.some(q=>!q.answered)} onClick={nextSection}>Continue →</button></>:<><h2>Declaration & Signature</h2><p style={{lineHeight:1.7}}>I certify that I have received this site orientation, understand my responsibility to proceed safely, and will follow Southwest Construction Management site rules and applicable Nova Scotia safety requirements.</p><label style={{display:'flex',gap:8}}><input type="checkbox" checked={run.declaration} onChange={e=>setRun({...run,declaration:e.target.checked})}/> I agree to this declaration.</label><label style={{display:'block',marginTop:16}}>Type your full name as your electronic signature<input style={{...field,marginTop:8}} value={run.signature} onChange={e=>setRun({...run,signature:e.target.value})}/></label><p style={{color:UI.sub,fontSize:12}}>Completing the form does not approve the orientation; the safety manager must verify it.</p><button style={btn} onClick={completeOrientation} disabled={busy||!run.declaration||!run.signature.trim()}>Submit for Safety Manager Review</button></>}</div>}
</main></div>;
}
function WorkerDetails({worker,attempts,onEdit,onOrient,onTicket,onVerify,onHistory,onWorkerLink,onDelete,busy,onClose}){const [ticket,setTicket]=useState({type:'',issued_date:'',expiry_date:'',holder_name:'',notes:''}),[file,setFile]=useState(null),[date,setDate]=useState(''),[reviewer,setReviewer]=useState(''),[scan,setScan]=useState(null);const attempt=attempts.slice().sort((a,b)=>String(b.started_at).localeCompare(String(a.started_at)))[0];const readFile=f=>new Promise((resolve,reject)=>{if(!f)return resolve(null);if(f.size>2200000)return reject(Error('File exceeds 2.2 MB'));const reader=new FileReader();reader.onload=()=>resolve({name:f.name,type:f.type,data:reader.result});reader.onerror=reject;reader.readAsDataURL(f)});return <div style={{...card,marginTop:16,borderColor:UI.teal}}><div style={{display:'flex',justifyContent:'space-between'}}><h3>{worker.name} · {worker.company}</h3><button style={secondary} onClick={onClose}>Close</button></div><p style={{color:UI.sub}}>{worker.trade} · {worker.occupation||'Occupation not provided'} · Language: {worker.preferred_language||'en'}</p><div style={{display:'flex',gap:10,flexWrap:'wrap',marginBottom:18}}>{onOrient&&<button style={btn} onClick={onOrient}>Start Digital Orientation</button>}{onEdit&&<button style={secondary} onClick={onEdit}>Edit Profile</button>}{onWorkerLink&&<button style={secondary} onClick={onWorkerLink}>Copy secure return link</button>}{onDelete&&<button disabled={busy} onClick={onDelete} style={{...secondary,border:'1px solid #E26B6B',color:'#FFB8B8'}}>Delete Worker</button>}</div><h4>Orientation</h4><p>{attempt?.approval_status==='approved'?'✓ Safety manager verified':attempt?.status==='completed'?'Awaiting safety manager approval':worker.historical_orientation?`Previously oriented: ${worker.historical_orientation.date}`:'No completed orientation'}</p>{onHistory&&<details><summary>Add existing signed paper orientation</summary><div style={{display:'flex',gap:10,flexWrap:'wrap',marginTop:12}}><input type="date" value={date} onChange={e=>setDate(e.target.value)} style={field}/><input placeholder="Name of Southwest verifier" value={reviewer} onChange={e=>setReviewer(e.target.value)} style={field}/><input type="file" accept="image/*,application/pdf" onChange={e=>setScan(e.target.files[0])}/><button style={btn} onClick={async()=>{try{await onHistory(date,reviewer,await readFile(scan))}catch(e){alert(e.message)}}}>Save Historical Record</button></div></details>}<h4>Tickets ({(worker.tickets||[]).length})</h4>{(worker.tickets||[]).map(t=><div key={t.id} style={{padding:11,marginBottom:8,background:UI.bg,borderRadius:8}}><b>{t.type}</b> · {t.expiry_date||'No expiry'} · {t.status||'needs_review'} {t.attachment&&<a href={t.attachment.data} download={t.attachment.name} style={{color:UI.teal,marginLeft:9}}>View file</a>}{onVerify&&t.status!=='verified'&&<button style={{...secondary,marginLeft:12}} onClick={()=>onVerify(t.id)}>Verify ticket</button>}</div>)}{onTicket&&<details><summary>Add a certificate / renewal</summary><div style={{display:'grid',gap:9,marginTop:10}}>{[['Certificate type (optional)','type'],['Name on certificate (optional)','holder_name'],['Issue date (optional)','issued_date'],['Expiry date (optional)','expiry_date'],['Notes','notes']].map(([name,k])=><label key={k}>{name}<input type={k.includes('date')?'date':'text'} value={ticket[k]} onChange={e=>setTicket({...ticket,[k]:e.target.value})} style={field}/></label>)}<input type="file" accept="image/*,application/pdf" onChange={e=>setFile(e.target.files[0])}/><button style={btn} disabled={!file} onClick={async()=>{if(await onTicket(ticket,file)){setFile(null);setTicket({type:'',issued_date:'',expiry_date:'',holder_name:'',notes:''})}}}>Add Ticket</button></div></details>}</div>}

function CertificateReview({worker,ticket,busy,manager,onVerify,onViewOriginal}){
const [details,setDetails]=useState({type:ticket.type||'',holder_name:ticket.holder_name||worker.name||'',issued_date:ticket.issued_date||'',expiry_date:ticket.expiry_date||''});
const mismatch=details.holder_name.trim()&&cleanName(details.holder_name)!==cleanName(worker.name);
return <section style={{...card,marginBottom:14}}><div style={{display:'flex',justifyContent:'space-between',gap:12,flexWrap:'wrap'}}><div><h3 style={{margin:0}}>{worker.name} · {worker.company}</h3><p style={{color:UI.sub,margin:'6px 0'}}>{worker.trade} · {ticket.type||'Unidentified certificate'} · Pending review</p></div>{ticket.attachment?.data&&<a style={{...secondary,textDecoration:'none'}} href={ticket.attachment.data} download={ticket.attachment.name||'certificate'}>View original certificate</a>}{ticket.attachment?.storage_path&&<button style={secondary} onClick={()=>onViewOriginal(worker.id,ticket.id)}>View original certificate</button>}</div><div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(190px,1fr))',gap:12,marginTop:12}}>{[['Certificate type','type'],['Name printed on certificate','holder_name'],['Issue date (optional)','issued_date'],['Expiry date','expiry_date']].map(([label,k])=><label key={k} style={{fontSize:12,color:UI.sub}}>{label}<input style={{...field,marginTop:5}} type={k.includes('date')?'date':'text'} value={details[k]} readOnly={!manager} onChange={e=>manager&&setDetails({...details,[k]:e.target.value})}/></label>)}</div>{mismatch&&<p style={{color:UI.yellow}}>Name on certificate differs from registered worker — check the original document.</p>}{manager&&<button style={{...btn,marginTop:14}} disabled={busy||!details.type.trim()||!details.holder_name.trim()||!details.expiry_date} onClick={()=>onVerify(details)}>Verify certificate</button>}{!manager&&<p style={{fontSize:12,color:UI.sub}}>Read-only staff access. The safety manager reviews and verifies certificates.</p>}</section>
}
