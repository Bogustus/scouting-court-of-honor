import type { ScoutAwardInfo } from "./parser";
import { Document, Paragraph, TextRun, Packer, HeadingLevel, AlignmentType } from "docx";

export interface ScriptOptions {
  date: string;
  time: string;
  location: string;
  mc1Name: string;
  mc2Name: string;
  scoutmasterName: string;
  introTitle: string;
  colorGuardNames: string;
}

const RANK_ORDER = [
  "Scout",
  "Tenderfoot",
  "Second Class",
  "First Class",
  "Star",
  "Life",
  "Eagle",
];

export async function generateDocxScript(data: ScoutAwardInfo[], opts: ScriptOptions) {
  const { ranks, meritBadges, awards } = groupData(data);

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          new Paragraph({
            text: "Court of Honor Ceremony",
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
          }),
          new Paragraph({
            text: opts.date,
            heading: HeadingLevel.HEADING_2,
            alignment: AlignmentType.CENTER,
          }),
          new Paragraph({ text: `Date: ${opts.date} ${opts.time}` }),
          new Paragraph({ text: `Location: ${opts.location}` }),
          new Paragraph({
            text: `Masters of Ceremony: ${opts.mc1Name} and ${opts.mc2Name}`,
          }),
          new Paragraph({
            text: `Color Guard: ${opts.colorGuardNames}`,
          }),
          new Paragraph(" "),

          new Paragraph({
            children: [
              new TextRun({ text: `${opts.mc1Name} (Call to order): `, bold: true }),
              new TextRun(
                "Ladies & Gentlemen, please find your seats as we'll be starting in a moment."
              ),
            ],
          }),
          new Paragraph(" "),

          // Standard Opening
          new Paragraph({
            children: [
              new TextRun({ text: `${opts.mc2Name}: `, bold: true }),
              new TextRun(
                `Hello, I am ${opts.mc2Name}. ${opts.mc1Name} and I will be your MCs for tonight's program.`
              ),
            ],
          }),
          new Paragraph("Will the troop & audience please rise, hats off."),
          new Paragraph("Color guard, advance. [wait for color guard to reach the front and stop]"),
          new Paragraph("Prepare to post the colors. [wait for color guard to cross flags and arrive at flag stands and stop]"),
          new Paragraph("Post the colors."),
          new Paragraph("Scouts, salute."),
          new Paragraph('Please join me in the Pledge of Allegiance. [“I pledge allegiance to the Flag …"]'),
          new Paragraph('Scout sign, The Scout Oath.  [“On my honor I will do my best…]'),
          new Paragraph('The Scout Law [“A scout is…”]'),
          new Paragraph("Two."),
          new Paragraph("Color guard, return to ranks. [Wait for color guard to re-form as a single unit and stop]"),
          new Paragraph("Color guard about, face."),
          new Paragraph("Color guard forward march."),
          new Paragraph("Color guard, dismissed. Troop at ease"),
          new Paragraph(" "),

          new Paragraph({
            children: [
              new TextRun({ text: `${opts.mc1Name}: `, bold: true }),
              new TextRun(
                `At this time I would like to introduce our ${opts.introTitle} to give our ceremony's introduction.`
              ),
            ],
          }),
          new Paragraph(`[ ${opts.scoutmasterName}: Introduction ]`),
          new Paragraph(" "),

          // Candle Ceremony
          new Paragraph({
            text: "Candle Ceremony",
            heading: HeadingLevel.HEADING_2,
          }),
          new Paragraph({
            children: [
              new TextRun({ text: `${opts.mc1Name}: `, bold: true }),
              new TextRun("We are going to begin by lighting this candle which represents the spirit of Scouting."),
            ],
          }),
          new Paragraph("We will now light a candle for each point of the Scout Law."),
          new Paragraph("A Scout is:"),
          new Paragraph(" "),
          new Paragraph({
            children: [
              new TextRun({ text: "TRUSTWORTHY: ", bold: true }),
              new TextRun("A Scout tells the truth. They keep their promises. Honesty is part of the code of conduct. People can depend on scouts."),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "LOYAL: ", bold: true }),
              new TextRun("Scouts are true to their family, Scout leaders, friends, school, and nation."),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "HELPFUL: ", bold: true }),
              new TextRun("Scouts are concerned about other people. They do things willingly for others without pay or reward."),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "FRIENDLY: ", bold: true }),
              new TextRun("Scouts are friends to all. They are kind to other Scouts. They seek to understand others. They respect those with ideas and customs other than their own."),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "COURTEOUS: ", bold: true }),
              new TextRun("Scouts are polite to everyone regardless of age or position. They know good manners make it easier for people to get along together."),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "KIND: ", bold: true }),
              new TextRun("Scouts understand there is strength in being gentle. They treat others as they want to be treated. They do not hurt or kill harmless things without reason."),
            ],
          }),
          new Paragraph(" "),
          new Paragraph({
            children: [
              new TextRun({ text: `${opts.mc2Name}: `, bold: true }),
              new TextRun({ text: "OBEDIENT: ", bold: true }),
              new TextRun("Scouts follow the rules of their family, school, and troop. They obey the laws of their community and country. If they think these rules and laws are unfair, they try to have them changed in an orderly manner rather than disobey them."),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "CHEERFUL: ", bold: true }),
              new TextRun("Scouts look for the bright side of things. They cheerfully do tasks that come their way. They try to make others happy."),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "THRIFTY: ", bold: true }),
              new TextRun("Scouts work to pay their way and to help others. They save for unforeseen needs. They protect and conserve natural resources. They carefully use time and property."),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "BRAVE: ", bold: true }),
              new TextRun("Scouts can face danger even if they are afraid. They have the courage to stand for what they think is right even if others laugh at or threaten them."),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "CLEAN: ", bold: true }),
              new TextRun("Scouts keep their body and mind fit and clean. They go around with those who believe in living by these same ideals. They help keep their home and community clean."),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "REVERENT: ", bold: true }),
              new TextRun("Scouts are reverent toward God. They are faithful in their religious duties. They respect the beliefs of others."),
            ],
          }),
          new Paragraph(" "),

          // Rank Advancements
          new Paragraph({
            text: "Award Rank Advancement",
            heading: HeadingLevel.HEADING_2,
          }),
          new Paragraph({
            children: [
              new TextRun({ text: `${opts.mc1Name}: `, bold: true }),
              new TextRun(
                "Rank Advancement is an important part of the Scouting program. It gives the Scout opportunities to learn new skills and have new adventures. Presenting the Scout with a new rank lets them show off the completion of a set of skills and adventures, and hopefully, encourages them on the Trail of Eagle Scout."
              ),
            ],
          }),
          new Paragraph("At this time, we would like to recognize those Scouts who have earned Rank Advancements."),
          new Paragraph("Scouting America recognizes your achievements by awarding badges of rank. There are seven ranks: Scout, Tenderfoot, Second Class, First Class, Star, Life and the highest rank is the coveted “Eagle.”"),
          new Paragraph("Some examples of Famous Eagle Scouts:"),
          new Paragraph({ text: "James Valentine the guitarist of the band Maroon 5", bullet: { level: 0 } }),
          new Paragraph({ text: "Dan Reynolds lead singer of the band Imagine Dragons", bullet: { level: 0 } }),
          new Paragraph({ text: "Glen Schofield one of the creators of the call of duty franchise", bullet: { level: 0 } }),
          new Paragraph({
            children: [
              new TextRun({ text: `${opts.mc2Name}: `, bold: true }),
              new TextRun(
                "Scouts, after your name is called, please come up and collect your award and line up on the front of the stage. Audience, please hold your applause until all the scouts are called for each rank."
              ),
            ],
          }),
          new Paragraph(" "),

          ...generateRankSection(ranks, opts),

          // Special Awards
          ...(awards.length > 0 ? [
            new Paragraph({
              text: "Awards",
              heading: HeadingLevel.HEADING_2,
            }),
            new Paragraph({
              children: [
                new TextRun({ text: `${opts.mc2Name}: `, bold: true }),
                new TextRun(
                  "Occasionally, scouts have the opportunity to earn special awards at camp, through their faith, or during troop activities. We are proud to present these special awards."
                ),
              ],
            }),
            new Paragraph("[MC Note: Call forward each scout with an award.]"),
            new Paragraph(" "),
            ...generateAwardsSection(awards, opts),
          ] : []),

          // Merit Badges
          ...(meritBadges.length > 0 ? [
            new Paragraph({
              text: "Award Merit Badges, BSA Awards, and Special Recognition",
              heading: HeadingLevel.HEADING_2,
            }),
            new Paragraph({
              children: [
                new TextRun({ text: `${opts.mc1Name}: `, bold: true }),
                new TextRun(
                  "At this time, we would like to present the Merit Badges. Scouts please come up and collect your merit badges when I call your name. Audience, please hold your applause to the end."
                ),
              ],
            }),
            new Paragraph(" "),
            ...generateMeritBadgeSection(meritBadges, opts),
          ] : []),

          // Closing
          new Paragraph(" "),
          new Paragraph({
            children: [
              new TextRun({ text: `${opts.mc2Name}: `, bold: true }),
              new TextRun(
                `That concludes our presentation of merit badges and BSA awards. ${opts.scoutmasterName} has some closing remarks before we have Flags.`
              ),
            ],
          }),
          new Paragraph(" "),
          new Paragraph(`[ ${opts.scoutmasterName}: Closing remarks ]`),
          new Paragraph(" "),
          new Paragraph({
            children: [
              new TextRun({ text: `${opts.mc1Name}: `, bold: true }),
              new TextRun(`Thank you. Will the troop & audience please rise, hats off.`),
            ],
          }),
          new Paragraph("Color guard, forward march. [wait for color guard to reach the front and stop]"),
          new Paragraph("Color guard, prepare to retire the colors. [wait for color guard to arrive at flag stands and stop]"),
          new Paragraph("Scout salute. [salute for about 3 seconds]"),
          new Paragraph("Two."),
          new Paragraph("Color guard, retrieve the colors."),
          new Paragraph("Color guard, return to ranks. [wait for color guard to re-form as a single unit and stop]"),
          new Paragraph("Color guard about, face"),
          new Paragraph("Color guard, forward march. [wait for color guard to reach the back of the room]"),
          new Paragraph("Color guard, dismissed."),
          new Paragraph("Troop dismissed. Thank you for coming"),
        ],
      },
    ],
  });

  return Packer.toBlob(doc);
}

function groupData(data: ScoutAwardInfo[]) {
  const ranks: { [rank: string]: string[] } = {
    Scout: [],
    Tenderfoot: [],
    "Second Class": [],
    "First Class": [],
    Star: [],
    Life: [],
    Eagle: [],
  };

  const meritBadges: { [scout: string]: string[] } = {};
  const awards: { scout: string; award: string }[] = [];

  for (const item of data) {
    const rawType = item.type.toLowerCase();

    if (rawType.includes("rank")) {
      const matchRank = RANK_ORDER.find((r) => r.toLowerCase() === item.award.toLowerCase());
      if (matchRank) {
        if (!ranks[matchRank].includes(item.scoutName)) {
          ranks[matchRank].push(item.scoutName);
        }
      } else {
        // Fallback or Unknown rank
        if (!ranks[item.award]) {
          ranks[item.award] = [];
        }
        if (!ranks[item.award].includes(item.scoutName)) {
          ranks[item.award].push(item.scoutName);
        }
      }
    } else if (rawType.includes("merit badge")) {
      if (!meritBadges[item.scoutName]) meritBadges[item.scoutName] = [];
      const cleanAward = item.award.replace("*", ""); // remove * commonly used for Eagle badges
      if (!meritBadges[item.scoutName].includes(cleanAward)) {
        meritBadges[item.scoutName].push(cleanAward);
      }
    } else {
      awards.push({ scout: item.scoutName, award: item.award });
    }
  }

  // Sort scout names in rank maps
  for (const group in ranks) {
    ranks[group].sort();
  }

  // Sort scouts alphabetically for merit badges
  const sortedMeritBadges = Object.keys(meritBadges)
    .sort()
    .map((name) => ({ name, badges: meritBadges[name].sort() }));

  return { ranks, meritBadges: sortedMeritBadges, awards };
}

const RANK_PROSE: { [key: string]: string } = {
  Scout: "The first award is Scout. To earn this award, a new Scout must agree to live by the Scout Oath and Law and complete a number of other assignments. I present to you the badge and rank advancement card, documenting your achievement, as well as the parent lapel pin.",
  Tenderfoot: "Tenderfoot requirements offer a taste of the great adventures awaiting you in Scouting and can give you the basic skills you'll need to begin taking part in those adventures. You may have met many challenges in earning the Tenderfoot badge and are to be congratulated.",
  "Second Class": "To earn Second Class, a Scout must learn how to use a map and compass, how and when to build a campfire, and to safely use pocket knives and wood tools. Second Class Scouts have proven their abilities in camping, first aid, and swimming, and other Scout skills.",
  "First Class": "The founder of Scouting, Lord Baden Powell, said that all Scouts should earn First Class. Now you have tested yourself even more. You have tried greater adventures and practiced your Scout skills many times. With your confidence and knowledge, you now have, people will expect more of you, and you will expect more of yourself. You are prepared to be more of a leader in your patrol, your troop, and your community.",
  Star: "As you earn your Star Rank, you have more freedom to choose the directions that interest you. The focus shifts from basic Scout skills to earning the first six merit badges you will need for Eagle. Requirements now include service to others. The Star Rank also requires the Scout to be active in his troop for at least four months. Of course, it may have taken longer. In addition, the Star Scout must serve his or her troop in a position of leadership for at least four months and take part in at least one service project.",
  Life: "The Life Rank is one of the rarest ranks. This is the last rank before Eagle. You could complete your Eagle Rank now in as few as 6 months. We congratulate you and encourage you to reach for that next step. You have earned more than half of the merit badges required for Eagle. The Life Rank also requires the Scout to be active in his or her troop for at least six months, serve his troop in a position of leadership for at least six months, and take part in at least one service project.",
  Eagle: "Eagle Scout is the highest rank attainable in the Scouts BSA program. The Eagle Scout must demonstrate Scout Spirit, an ideal attitude based upon the Scout Oath and Law, service, and leadership. Being an Eagle Scout is important because it requires immense hard work, dedication, and service to others."
};

function generateRankSection(
  ranks: { [rank: string]: string[] },
  opts: ScriptOptions
): Paragraph[] {
  const paras: Paragraph[] = [];
  let useMc1 = false; // toggle between MCs

  for (const r of RANK_ORDER) {
    const list = ranks[r];
    if (list && list.length > 0) {
      let mc = useMc1 ? opts.mc1Name : opts.mc2Name;
      useMc1 = !useMc1;

      // Handle conflict: If MC1 is getting the award, MC2 must present it.
      if (list.some(n => n.toLowerCase().includes(opts.mc1Name.toLowerCase()))) {
        mc = opts.mc2Name;
        useMc1 = true; // reset so next is mc1
      } else if (list.some(n => n.toLowerCase().includes(opts.mc2Name.toLowerCase()))) {
        mc = opts.mc1Name;
        useMc1 = false;
      }

      paras.push(
        new Paragraph({
          text: r.toUpperCase(),
          heading: HeadingLevel.HEADING_3,
        })
      );

      // Add the descriptive prose for the rank
      if (r === "Eagle") {
        paras.push(
          new Paragraph({
            children: [
              new TextRun({ text: `${mc}: `, bold: true }),
              new TextRun("Would any and all Eagle Scouts in the room please rise? [Pause]"),
            ]
          }),
          new Paragraph("You may be seated."),
          new Paragraph("The Eagle Rank is Scouting’s highest award. Only a small percentage of Scouts have ever reached this lofty goal. The Eagle Rank is presented in a special Eagle Scout Court of Honor."),
          new Paragraph("Tonight we are recognizing scouts who are reaching Eagle Rank. Would you please stand and be recognized for your achievement.")
        );

        for (const s of list) {
          paras.push(new Paragraph({ text: s, bullet: { level: 0 } }));
        }

        paras.push(new Paragraph(" "));
        paras.push(new Paragraph("[Lead applause]"));
        paras.push(new Paragraph(" "));
      } else {
        if (RANK_PROSE[r]) {
          paras.push(
            new Paragraph({
              children: [
                new TextRun({ text: `${mc}: `, bold: true }),
                new TextRun(RANK_PROSE[r]),
              ],
            })
          );
          paras.push(new Paragraph(" "));
        }

        paras.push(
          new Paragraph({
            children: [
              new TextRun({ text: `${mc}: `, bold: true }),
              new TextRun(`The following Scouts have achieved ${r} Rank:`),
            ],
          })
        );

        for (const s of list) {
          paras.push(new Paragraph({ text: s, bullet: { level: 0 } }));
        }

        paras.push(new Paragraph(" "));
        paras.push(new Paragraph(`[Wait for all scouts to be in place on the stage]`));
        paras.push(
          new Paragraph(`New ${r.toUpperCase()} scouts, wear your badge with pride. [Lead applause]`)
        );
        paras.push(new Paragraph(" "));
      }
    }
  }

  return paras;
}

function generateAwardsSection(awards: { scout: string; award: string }[], opts: ScriptOptions): Paragraph[] {
  const paras: Paragraph[] = [];
  let lastMc = opts.mc2Name; // Assume MC2 just finished the intro

  for (const { scout, award } of awards) {
    let mc = opts.mc2Name;

    if (scout.toLowerCase().includes(opts.mc2Name.toLowerCase())) {
      mc = opts.mc1Name;
    }

    if (mc !== lastMc) {
      paras.push(new Paragraph(" "));
      paras.push(
        new Paragraph({
          children: [new TextRun({ text: `${mc}: `, bold: true })],
        })
      );
      lastMc = mc;
    }

    paras.push(
      new Paragraph({
        children: [
          new TextRun({ text: `${scout}`, bold: true }),
          new TextRun(` gets the ${award}.`),
        ],
      })
    );
  }
  paras.push(new Paragraph(" "));

  return paras;
}

function generateMeritBadgeSection(meritBadges: { name: string; badges: string[] }[], opts: ScriptOptions): Paragraph[] {
  const paras: Paragraph[] = [];
  let useMc1 = true;
  const isLargeSet = meritBadges.length > 20;
  const halfPoint = Math.ceil(meritBadges.length / 2);
  let lastMc = opts.mc1Name; // MC1 reads the intro paragraph right before this

  for (let i = 0; i < meritBadges.length; i++) {
    const { name, badges } = meritBadges[i];
    let badgeText = "";
    if (badges.length === 1) {
      badgeText = `the ${badges[0]} merit badge`;
    } else if (badges.length === 2) {
      badgeText = `the ${badges[0]} and ${badges[1]} merit badges`;
    } else {
      const allButLast = badges.slice(0, badges.length - 1).join(", ");
      const last = badges[badges.length - 1];
      badgeText = `the ${allButLast}, and ${last} merit badges`;
    }

    let mc: string;
    
    if (isLargeSet) {
      mc = i < halfPoint ? opts.mc1Name : opts.mc2Name;
    } else {
      mc = useMc1 ? opts.mc1Name : opts.mc2Name;
      useMc1 = !useMc1;
    }

    // Still override if the assigned MC is reading their own name
    if (name.toLowerCase().includes(opts.mc1Name.toLowerCase())) {
      mc = opts.mc2Name;
      useMc1 = true; // For alternating, reset the tracking so next is MC1
    } else if (name.toLowerCase().includes(opts.mc2Name.toLowerCase())) {
      mc = opts.mc1Name;
      useMc1 = false;
    }

    if (mc !== lastMc) {
      paras.push(new Paragraph(" "));
      paras.push(
        new Paragraph({
          children: [new TextRun({ text: `${mc}: `, bold: true })],
        })
      );
      lastMc = mc;
    }

    paras.push(
      new Paragraph({
        children: [
          new TextRun({ text: `${name}`, bold: true }),
          new TextRun(` has earned ${badgeText}.`),
        ],
      })
    );
  }

  paras.push(new Paragraph(" "));
  paras.push(new Paragraph("[Lead applause]"));
  paras.push(new Paragraph(" "));

  return paras;
}
