// ============================================================
// MALAYPRAG MINIMAL-PAIR PIPELINE
//
// Produces:
// 1. All Minimal Pairs
// 2. Full Triplets          = naturally occurring kan-ke-null
// 3. Summary
// 4. Completed Triplets     = kan-ke pairs + synthesised null
//
// Does NOT modify the original MalayPrag sheet.
// Does NOT create a Manual Check sheet.
// ============================================================


// ============================================================
// MAIN ONE-CLICK FUNCTION
// ============================================================

function runMalayPragPipeline() {
  findMalayPragMinimalPairs();
  createCompletedTriplets();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Minimal pairs and completed triplets generated.",
    "MalayPrag complete",
    5
  );
}


// ============================================================
// PART 1: FIND NATURAL MINIMAL PAIRS
// ============================================================

function findMalayPragMinimalPairs() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const REQUIRED_COLUMNS = [
    "Text",
    "Source",
    "Particle",
    "Sentence_Type",
    "Epistemic_Stance",
    "Listener_Agreement",
    "Emotion",
    "Question_Type",
    "Particle_Position"
  ];

  // These sheets must NEVER be treated as the original dataset.
  const OUTPUT_SHEETS = [
    "All Minimal Pairs",
    "Full Triplets",
    "Summary",
    "Completed Triplets",
    "Manual Check" // included only so an old copy isn't used as source
  ];


  // ============================================================
  // 1. FIND ORIGINAL MALAYPRAG SHEET
  // ============================================================

  let sourceSheet = null;

  for (const sheet of ss.getSheets()) {

    if (OUTPUT_SHEETS.includes(sheet.getName())) {
      continue;
    }

    const lastColumn = sheet.getLastColumn();

    if (lastColumn === 0) {
      continue;
    }

    const headers = sheet
      .getRange(1, 1, 1, lastColumn)
      .getValues()[0]
      .map(h => String(h).trim());

    const hasAllColumns = REQUIRED_COLUMNS.every(
      column => headers.includes(column)
    );

    if (hasAllColumns) {
      sourceSheet = sheet;
      break;
    }
  }

  if (!sourceSheet) {
    throw new Error(
      "Could not find the original MalayPrag sheet."
    );
  }

  Logger.log(
    "Using source sheet: " + sourceSheet.getName()
  );


  // ============================================================
  // 2. READ DATA
  // ============================================================

  const data = sourceSheet
    .getDataRange()
    .getValues();

  if (data.length < 2) {
    throw new Error(
      "The MalayPrag source sheet contains no data."
    );
  }

  const headers = data[0]
    .map(h => String(h).trim());

  const rows = data.slice(1);

  const columnIndex = {};

  headers.forEach((header, index) => {
    columnIndex[header] = index;
  });


  // ============================================================
  // 3. HELPER FUNCTIONS
  // ============================================================

  function escapeRegex(text) {
    return String(text)
      .replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }


  function normalizeParticle(value) {

    if (
      value === null ||
      value === undefined ||
      String(value).trim() === ""
    ) {
      return "null";
    }

    const p = String(value)
      .trim()
      .toLowerCase();

    if (
      [
        "null",
        "none",
        "nan",
        "ø",
        "∅",
        "na",
        "n/a"
      ].includes(p)
    ) {
      return "null";
    }

    if (p === "kan") return "kan";
    if (p === "ke") return "ke";

    return p;
  }


  function normalizePosition(value) {

    if (
      value === null ||
      value === undefined ||
      String(value).trim() === ""
    ) {
      return "null";
    }

    const p = String(value)
      .trim()
      .toLowerCase();

    if (
      [
        "initial",
        "sentence-initial",
        "sentence initial"
      ].includes(p)
    ) {
      return "initial";
    }

    if (
      [
        "medial",
        "middle",
        "sentence-medial",
        "sentence medial"
      ].includes(p)
    ) {
      return "medial";
    }

    if (
      [
        "final",
        "sentence-final",
        "sentence final"
      ].includes(p)
    ) {
      return "final";
    }

    return p;
  }


  // ------------------------------------------------------------
  // Create particle-stripped key
  //
  // e.g.
  //
  // "dia dah makan kan?" -> "dia dah makan"
  // "dia dah makan ke?"  -> "dia dah makan"
  // ------------------------------------------------------------

  function makePairKey(
    textValue,
    particle,
    position
  ) {

    let text = String(textValue || "")
      .trim()
      .toLowerCase();

    // Standardise whitespace
    text = text.replace(/\s+/g, " ");

    // Ignore sentence-final punctuation for matching
    text = text
      .replace(/[?!.,;:…]+$/g, "")
      .trim();


    // Null has nothing to remove
    if (particle === "null") {
      return text;
    }

    if (!["kan", "ke"].includes(particle)) {
      return text;
    }

    const escapedParticle =
      escapeRegex(particle);


    // INITIAL
    if (position === "initial") {

      const regex = new RegExp(
        "^\\s*" +
        escapedParticle +
        "\\b[\\s,]*",
        "i"
      );

      text = text.replace(regex, "");
    }


    // FINAL
    else if (position === "final") {

      const regex = new RegExp(
        "[\\s,]*\\b" +
        escapedParticle +
        "\\s*$",
        "i"
      );

      text = text.replace(regex, "");
    }


    // MEDIAL
    else if (position === "medial") {

      const regex = new RegExp(
        "\\b" +
        escapedParticle +
        "\\b",
        "i"
      );

      text = text.replace(regex, "");
    }


    // FALLBACK
    else {

      const finalRegex = new RegExp(
        "[\\s,]*\\b" +
        escapedParticle +
        "\\s*$",
        "i"
      );

      if (finalRegex.test(text)) {

        text = text.replace(
          finalRegex,
          ""
        );

      } else {

        const generalRegex =
          new RegExp(
            "\\b" +
            escapedParticle +
            "\\b",
            "i"
          );

        text = text.replace(
          generalRegex,
          ""
        );
      }
    }


    text = text
      .replace(/\s+/g, " ")
      .trim();

    return text;
  }


  function countParticleTokens(
    text,
    particle
  ) {

    if (!["kan", "ke"].includes(particle)) {
      return 0;
    }

    const regex = new RegExp(
      "\\b" +
      escapeRegex(particle) +
      "\\b",
      "gi"
    );

    const matches =
      String(text || "").match(regex);

    return matches
      ? matches.length
      : 0;
  }


  function getPairType(particles) {

    const set = new Set(particles);

    const hasKan = set.has("kan");
    const hasKe = set.has("ke");
    const hasNull = set.has("null");

    if (
      hasKan &&
      hasKe &&
      hasNull
    ) {
      return "kan-ke-null";
    }

    if (
      hasKan &&
      hasKe
    ) {
      return "kan-ke";
    }

    if (
      hasKan &&
      hasNull
    ) {
      return "kan-null";
    }

    if (
      hasKe &&
      hasNull
    ) {
      return "ke-null";
    }

    return "other";
  }


  function particleOrder(particle) {

    const order = {
      "kan": 1,
      "ke": 2,
      "null": 3
    };

    return order[particle] || 99;
  }


  // ============================================================
  // 4. PROCESS EVERY ROW
  // ============================================================

  const processedRows = [];

  rows.forEach((row, index) => {

    const text =
      row[columnIndex["Text"]];

    if (
      text === "" ||
      text === null ||
      text === undefined
    ) {
      return;
    }

    const particleOriginal =
      row[columnIndex["Particle"]];

    const positionOriginal =
      row[columnIndex["Particle_Position"]];

    const particle =
      normalizeParticle(
        particleOriginal
      );

    const position =
      normalizePosition(
        positionOriginal
      );

    const pairKey =
      makePairKey(
        text,
        particle,
        position
      );

    const tokenCount =
      countParticleTokens(
        text,
        particle
      );

    processedRows.push({

      originalRow: row,

      originalRowNumber:
        index + 2,

      particle:
        particle,

      position:
        position,

      pairKey:
        pairKey,

      tokenCount:
        tokenCount,

      needsManualCheck:
        tokenCount > 1

    });
  });


  // ============================================================
  // 5. GROUP BY PAIR KEY
  // ============================================================

  const groups = {};

  processedRows.forEach(item => {

    if (!groups[item.pairKey]) {
      groups[item.pairKey] = [];
    }

    groups[item.pairKey]
      .push(item);
  });


  // ============================================================
  // 6. KEEP GROUPS WITH AT LEAST TWO PARTICLE TYPES
  // ============================================================

  const validGroups = {};

  Object.keys(groups)
    .forEach(key => {

      const particles = [
        ...new Set(
          groups[key]
            .map(item =>
              item.particle
            )
        )
      ];

      if (particles.length >= 2) {
        validGroups[key] =
          groups[key];
      }
    });


  // ============================================================
  // 7. ASSIGN PAIR IDS
  // ============================================================

  const sortedKeys =
    Object.keys(validGroups)
      .sort();

  const pairIdMap = {};

  sortedKeys.forEach(
    (key, index) => {

      pairIdMap[key] =
        "MP" +
        String(index + 1)
          .padStart(3, "0");
    }
  );


  // ============================================================
  // 8. BUILD OUTPUT
  // ============================================================

  const output = [];

  sortedKeys.forEach(key => {

    const group =
      validGroups[key];

    const particles = [
      ...new Set(
        group.map(item =>
          item.particle
        )
      )
    ];

    const pairType =
      getPairType(
        particles
      );

    const particleSet =
      particles
        .sort(
          (a, b) =>
            particleOrder(a) -
            particleOrder(b)
        )
        .join(", ");

    group.forEach(item => {

      output.push({

        pairID:
          pairIdMap[key],

        pairKey:
          key,

        pairType:
          pairType,

        nParticleTypes:
          particles.length,

        particleSet:
          particleSet,

        particleNormalized:
          item.particle,

        positionNormalized:
          item.position,

        particleTokenCount:
          item.tokenCount,

        needsManualCheck:
          item.needsManualCheck,

        originalRow:
          item.originalRow

      });
    });
  });


  // ============================================================
  // 9. SORT KAN -> KE -> NULL
  // ============================================================

  output.sort(
    (a, b) => {

      if (
        a.pairID !== b.pairID
      ) {
        return a.pairID
          .localeCompare(
            b.pairID
          );
      }

      return (
        particleOrder(
          a.particleNormalized
        ) -
        particleOrder(
          b.particleNormalized
        )
      );
    }
  );


  // ============================================================
  // 10. OUTPUT HEADERS
  // ============================================================

  const outputHeaders = [

    "Pair_ID",
    "Pair_Key",
    "Pair_Type",
    "N_Particle_Types",
    "Particle_Set",

    ...headers,

    "Particle_Normalized",
    "Position_Normalized",
    "Particle_Token_Count",
    "Needs_Manual_Check"

  ];


  // ============================================================
  // 11. ALL MINIMAL-PAIR ROWS
  // ============================================================

  const allPairRows =
    output.map(item => [

      item.pairID,
      item.pairKey,
      item.pairType,
      item.nParticleTypes,
      item.particleSet,

      ...item.originalRow,

      item.particleNormalized,
      item.positionNormalized,
      item.particleTokenCount,
      item.needsManualCheck

    ]);


  // ============================================================
  // 12. NATURALLY OCCURRING FULL TRIPLETS
  // ============================================================

  const tripletRows =
    output
      .filter(
        item =>
          item.pairType ===
          "kan-ke-null"
      )
      .map(item => [

        item.pairID,
        item.pairKey,
        item.pairType,
        item.nParticleTypes,
        item.particleSet,

        ...item.originalRow,

        item.particleNormalized,
        item.positionNormalized,
        item.particleTokenCount,
        item.needsManualCheck

      ]);


  // ============================================================
  // 13. SUMMARY
  // ============================================================

  const summaryCounts = {};

  sortedKeys.forEach(key => {

    const particles = [
      ...new Set(
        validGroups[key]
          .map(item =>
            item.particle
          )
      )
    ];

    const type =
      getPairType(
        particles
      );

    summaryCounts[type] =
      (summaryCounts[type] || 0)
      + 1;
  });


  const summaryRows = [
    [
      "Pair_Type",
      "Number_of_Groups"
    ]
  ];


  [
    "kan-ke-null",
    "kan-ke",
    "kan-null",
    "ke-null",
    "other"
  ].forEach(type => {

    if (summaryCounts[type]) {

      summaryRows.push([
        type,
        summaryCounts[type]
      ]);
    }
  });


  summaryRows.push([
    "",
    ""
  ]);


  summaryRows.push([
    "Total minimal-pair groups",
    sortedKeys.length
  ]);


  summaryRows.push([
    "Naturally occurring full triplets",
    summaryCounts["kan-ke-null"] || 0
  ]);


  // ============================================================
  // 14. FUNCTION TO WRITE A SHEET
  // ============================================================

  function writeSheet(
    sheetName,
    headersArray,
    rowsArray
  ) {

    let sheet =
      ss.getSheetByName(
        sheetName
      );

    if (!sheet) {
      sheet =
        ss.insertSheet(
          sheetName
        );
    }

    sheet.clearContents();

    sheet
      .getRange(
        1,
        1,
        1,
        headersArray.length
      )
      .setValues(
        [headersArray]
      );


    if (rowsArray.length > 0) {

      sheet
        .getRange(
          2,
          1,
          rowsArray.length,
          headersArray.length
        )
        .setValues(
          rowsArray
        );
    }

    sheet.setFrozenRows(1);
  }


  // ============================================================
  // 15. WRITE OUTPUT SHEETS
  // ============================================================

  writeSheet(
    "All Minimal Pairs",
    outputHeaders,
    allPairRows
  );


  writeSheet(
    "Full Triplets",
    outputHeaders,
    tripletRows
  );


  // ============================================================
  // DELETE OLD MANUAL CHECK TAB IF IT EXISTS
  // ============================================================

  const oldManual =
    ss.getSheetByName(
      "Manual Check"
    );

  if (oldManual) {
    ss.deleteSheet(
      oldManual
    );
  }


  // ============================================================
  // WRITE SUMMARY
  // ============================================================

  let summarySheet =
    ss.getSheetByName(
      "Summary"
    );

  if (!summarySheet) {
    summarySheet =
      ss.insertSheet(
        "Summary"
      );
  }

  summarySheet.clearContents();

  summarySheet
    .getRange(
      1,
      1,
      summaryRows.length,
      2
    )
    .setValues(
      summaryRows
    );

  summarySheet.setFrozenRows(1);


  SpreadsheetApp.flush();

  Logger.log(
    "Found " +
    sortedKeys.length +
    " minimal-pair groups."
  );
}



// ============================================================
// PART 2: TURN KAN-KE PAIRS INTO KAN-KE-NULL TRIPLETS
// ============================================================

function createCompletedTriplets() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  // ============================================================
  // 1. GET MINIMAL-PAIR SHEET
  // ============================================================

  const sourceSheet =
    ss.getSheetByName(
      "All Minimal Pairs"
    );

  if (!sourceSheet) {

    throw new Error(
      'Could not find "All Minimal Pairs". Run findMalayPragMinimalPairs() first.'
    );
  }


  const data =
    sourceSheet
      .getDataRange()
      .getValues();


  if (data.length < 2) {

    throw new Error(
      "No minimal-pair data found."
    );
  }


  const headers =
    data[0].map(
      h => String(h).trim()
    );

  const rows =
    data.slice(1);


  // ============================================================
  // 2. COLUMN INDEXES
  // ============================================================

  const col = {};

  headers.forEach(
    (header, index) => {

      col[header] = index;

    }
  );


  const required = [

    "Pair_ID",
    "Pair_Key",
    "Pair_Type",
    "N_Particle_Types",
    "Particle_Set",
    "Text",
    "Particle"

  ];


  required.forEach(name => {

    if (!(name in col)) {

      throw new Error(
        "Missing required column: " +
        name
      );
    }
  });


  // ============================================================
  // 3. GROUP BY PAIR ID
  // ============================================================

  const groups = {};

  rows.forEach(row => {

    const pairID =
      String(
        row[col["Pair_ID"]]
      ).trim();

    if (!pairID) {
      return;
    }

    if (!groups[pairID]) {
      groups[pairID] = [];
    }

    groups[pairID]
      .push(row);
  });


  // ============================================================
  // 4. CREATE COMPLETE TRIPLETS
  //
  // Only groups containing BOTH kan and ke are used.
  //
  // If null already exists, it is preserved.
  // Otherwise a synthesised null row is added.
  // ============================================================

  const outputRows = [];


  Object.keys(groups)
    .sort()
    .forEach(pairID => {

      const group =
        groups[pairID];


      const particles =
        new Set(
          group.map(row =>
            String(
              row[col["Particle"]]
            )
              .trim()
              .toLowerCase()
          )
        );


      // We specifically want kan-ke(-null) sets.
      if (
        !particles.has("kan") ||
        !particles.has("ke")
      ) {
        return;
      }


      // --------------------------------------------------------
      // COPY EXISTING ROWS
      // --------------------------------------------------------

      group.forEach(
        originalRow => {

          const row =
            [...originalRow];

          // Completed-set metadata
          row[col["Pair_Type"]] =
            "kan-ke-null";

          row[col["N_Particle_Types"]] =
            3;

          row[col["Particle_Set"]] =
            "kan, ke, null";

          outputRows.push(row);
        }
      );


      // --------------------------------------------------------
      // CREATE NULL VERSION IF MISSING
      // --------------------------------------------------------

      if (!particles.has("null")) {

        const nullRow =
          new Array(
            headers.length
          ).fill("");


        const pairKey =
          String(
            group[0][
              col["Pair_Key"]
            ]
          ).trim();


        // Pair metadata
        nullRow[
          col["Pair_ID"]
        ] =
          pairID;

        nullRow[
          col["Pair_Key"]
        ] =
          pairKey;

        nullRow[
          col["Pair_Type"]
        ] =
          "kan-ke-null";

        nullRow[
          col["N_Particle_Types"]
        ] =
          3;

        nullRow[
          col["Particle_Set"]
        ] =
          "kan, ke, null";


        // ======================================================
        // NULL TEXT
        //
        // Pair_Key is already the particle-stripped utterance.
        // ======================================================

        nullRow[
          col["Text"]
        ] =
          pairKey;

        nullRow[
          col["Particle"]
        ] =
          "null";


        // ======================================================
        // METADATA
        // ======================================================

        if ("Source" in col) {

          nullRow[
            col["Source"]
          ] =
            "Synthesised";
        }


        if ("Sentence_Type" in col) {

          nullRow[
            col["Sentence_Type"]
          ] =
            "Synthesised (null)";
        }


        if ("Particle_Position" in col) {

          nullRow[
            col["Particle_Position"]
          ] =
            "null";
        }


        if (
          "Particle_Normalized"
          in col
        ) {

          nullRow[
            col["Particle_Normalized"]
          ] =
            "null";
        }


        if (
          "Position_Normalized"
          in col
        ) {

          nullRow[
            col["Position_Normalized"]
          ] =
            "null";
        }


        if (
          "Particle_Token_Count"
          in col
        ) {

          nullRow[
            col["Particle_Token_Count"]
          ] =
            0;
        }


        if (
          "Needs_Manual_Check"
          in col
        ) {

          nullRow[
            col["Needs_Manual_Check"]
          ] =
            false;
        }


        // ======================================================
        // DO NOT COPY HUMAN PRAGMATIC ANNOTATIONS
        // ======================================================

        const annotationsToBlank = [

          "Epistemic_Stance",
          "Listener_Agreement",
          "Emotion",
          "Question_Type"

        ];


        annotationsToBlank.forEach(
          name => {

            if (name in col) {

              nullRow[
                col[name]
              ] = "";
            }
          }
        );


        outputRows.push(
          nullRow
        );
      }
    });


  // ============================================================
  // 5. SORT:
  //
  // MP001 kan
  // MP001 ke
  // MP001 null
  // MP002 kan
  // MP002 ke
  // MP002 null
  // ============================================================

  const particleOrder = {

    "kan": 1,
    "ke": 2,
    "null": 3

  };


  outputRows.sort(
    (a, b) => {

      const pairA =
        String(
          a[col["Pair_ID"]]
        );

      const pairB =
        String(
          b[col["Pair_ID"]]
        );


      if (pairA !== pairB) {

        return pairA
          .localeCompare(pairB);
      }


      const particleA =
        String(
          a[col["Particle"]]
        )
          .trim()
          .toLowerCase();


      const particleB =
        String(
          b[col["Particle"]]
        )
          .trim()
          .toLowerCase();


      return (

        (particleOrder[
          particleA
        ] || 99)

        -

        (particleOrder[
          particleB
        ] || 99)

      );
    }
  );


  // ============================================================
  // 6. CREATE / REFRESH COMPLETED TRIPLETS SHEET
  // ============================================================

  const outputName =
    "Completed Triplets";


  let outputSheet =
    ss.getSheetByName(
      outputName
    );


  if (!outputSheet) {

    outputSheet =
      ss.insertSheet(
        outputName
      );
  }


  outputSheet.clearContents();


  // Headers
  outputSheet
    .getRange(
      1,
      1,
      1,
      headers.length
    )
    .setValues(
      [headers]
    );


  // Data
  if (outputRows.length > 0) {

    outputSheet
      .getRange(
        2,
        1,
        outputRows.length,
        headers.length
      )
      .setValues(
        outputRows
      );
  }


  outputSheet.setFrozenRows(1);

  SpreadsheetApp.flush();


  // ============================================================
  // 7. FINAL MESSAGE
  // ============================================================

  const completedIDs =
    new Set(
      outputRows.map(
        row =>
          row[col["Pair_ID"]]
      )
    );


  Logger.log(
    "Created " +
    completedIDs.size +
    " completed kan-ke-null triplets."
  );
}

function createModelInputSheet() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Completed Triplets");

  if (!sourceSheet) {
    throw new Error('Could not find "Completed Triplets".');
  }

  const data = sourceSheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const rows = data.slice(1);

  const pairIndex = headers.indexOf("Pair_ID");
  const textIndex = headers.indexOf("Text");
  const particleIndex = headers.indexOf("Particle");

  if (
    pairIndex === -1 ||
    textIndex === -1 ||
    particleIndex === -1
  ) {
    throw new Error(
      'Required columns "Pair_ID", "Text", or "Particle" are missing.'
    );
  }

  const groups = {};

  rows.forEach(row => {

    const pairID = String(row[pairIndex]).trim();
    const text = String(row[textIndex]).trim();
    const particle = String(row[particleIndex]).trim().toLowerCase();

    if (!pairID) return;

    if (!groups[pairID]) {
      groups[pairID] = {
        kan: "",
        ke: "",
        null: ""
      };
    }

    if (particle === "kan") {
      groups[pairID].kan = text;
    }

    else if (particle === "ke") {
      groups[pairID].ke = text;
    }

    else if (particle === "null") {
      groups[pairID].null = text;
    }

  });


  const output = [
    [
      "Pair_ID",
      "Kan_Text",
      "Ke_Text",
      "Null_Text"
    ]
  ];


  Object.keys(groups)
    .sort()
    .forEach(pairID => {

      output.push([
        pairID,
        groups[pairID].kan,
        groups[pairID].ke,
        groups[pairID].null
      ]);

    });


  const outputName = "Model Input";

  let outputSheet = ss.getSheetByName(outputName);

  if (!outputSheet) {
    outputSheet = ss.insertSheet(outputName);
  }

  outputSheet.clearContents();

  outputSheet
    .getRange(
      1,
      1,
      output.length,
      4
    )
    .setValues(output);

  outputSheet.setFrozenRows(1);

  SpreadsheetApp.flush();

  ss.toast(
    "Model Input sheet created.",
    "MalayPrag",
    5
  );
}
