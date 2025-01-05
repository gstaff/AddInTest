/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

// Gets the table in the doucment by first cell text.
export async function getTable(context, firstCellText) {
  await context.document.body.paragraphs.load();
  await context.sync();
  let table = null;
  context.document.body.paragraphs.items.forEach((p) => {
    if (p.text === firstCellText) {
      try {
        table = p.parentTable;
        return;
      } catch {
        table = null;
      }
    }
  });
  return table;
}

// Finds and replaces placeholder text with value anywhere in the document.
export async function replacePlaceholder(context, placeholder, value) {
  console.log(`Replacing ${placeholder} with ${value}`);
  const body = context.document.body;
  const paragraphs = body.paragraphs;
  await paragraphs.load();
  await context.sync();
  paragraphs.items.forEach((p) => {
    if (p.tableNestingLevel > 0) {
      return;
    }
    const originalText = p.text;
    const newText = originalText.replace(placeholder, value);
    p.clear();
    p.insertText(newText, Word.InsertLocation.start);
  });
  await context.sync();
}

// Extracts a map of key, value pairs from a 3 column table "Tag", "Value", "Notes".
// Tags must have format [<TagName>].
export async function getValueMapping(context) {
  context.document.body.tables.load();
  await context.sync();
  const table = await getTable(context, "Tag");
  table.load(["values"]);
  await context.sync();

  const elements = table.values;
  const map = new Map();
  for (let i = 0; i < elements.length; i += 1) {
    const row = elements[i];
    const key = row[0];
    if (!key.startsWith("[")) {
      continue;
    }
    let value = row[1];
    if (!value) {
      value = "";
    }
    map.set(key.trim(), value.trim());
  }
  // table.delete();
  return map;
}

// Example of how we can fetch API data.
export async function getAPIData() {
  // Define the API URL
  const apiUrl = "https://pokeapi.co/api/v2/pokemon/ditto";
  const response = await fetch(apiUrl);
  return await response.json();
}

// Inserts an image to replace @[<tag>].
export async function insertImage(base64Image, tag) {
  return Word.run(async (context) => {
    await context.document.body.paragraphs.load();
    await context.sync();
    const p = context.document.body.paragraphs.items.filter((p) => p.text === tag)[0];
    if (p) {
      p.insertParagraph("", "After").insertInlinePictureFromBase64(
        base64Image.replace("data:image/png;base64,", ""),
        "End"
      );
      p.delete();
    }
    await context.sync();
  });
}

// Creates and adds a chart to the document.
export async function createChart(tag, clientValue, ranges) {
  console.log(`Creating chart for ${tag}`);
  // Get the div element
  const div = document.getElementById("renderDiv");

  // Create a new canvas element
  const canvas = document.createElement("canvas");

  // Set any attributes for the canvas (optional)
  canvas.width = 800;
  canvas.height = 100;

  // Append the canvas to the div
  div.appendChild(canvas);

  const metricName = tag.slice(1);

  const datasets = ranges.map((r) => {
    return { type: "bar", label: r.label, data: r.data, indexAxis: "y", backgroundColor: r.backgroundColor };
  });
  datasets.push({
    type: "line",
    label: `Your Value: ${clientValue}`,
    data: [clientValue],
    indexAxis: "y",
    backgroundColor: "blue",
  });

  // Min? max? title?
  const myChart = new Chart(canvas, {
    data: {
      datasets: datasets,
      labels: [metricName],
    },
    options: {
      scales: {
        x: {
          stacked: true,
        },
        y: {
          stacked: true,
        },
      },
      animation: {
        duration: 10,
        onComplete: function (chart) {
          if (chart.initial) {
            insertImage(myChart.toBase64Image(), tag);
            div.removeChild(canvas);
          }
        },
      },
    },
  });
}

// Deletes a section.
// Section format is #<SectionName>#\n<SectionContent.\n#<SectionName>#
export async function deleteSection(sectionName) {
  return Word.run(async (context) => {
    // Note: [^13] also removes the paragraph marker before so nothing is left.
    const searchQuery = `[^13]${sectionName}*${sectionName}`;

    // Queue a command to search the document with a wildcard
    const searchResults = context.document.body.search(searchQuery, { matchWildcards: true });

    // Queue a command to load the search results.
    searchResults.load();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log("Found count: " + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
      searchResults.items[i].delete();
    }

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    // await context.sync();
  });
}

export async function getRangesForTag(context, searchTag) {
  const colorMap = {
    Red: "hsla(0, 100%, 50%, 0.2)",
    Yellow: "hsla(39, 100%, 50%, 0.2)",
    Green: "hsla(147, 50%, 47%, 0.2)",
  };
  const table = await getTable(context, "Section Name");
  table.load(["values"]);
  await context.sync();
  const elements = table.values;
  const ranges = [];
  // TODO: Order ranges.
  for (let i = 0; i < elements.length; i += 1) {
    const row = elements[i];
    // const section = row[0].trim();
    const tag = row[1].trim();
    if (tag !== searchTag) {
      continue;
    }
    const low = row[2].trim();
    const high = row[3].trim();
    const label = row[4].trim();
    const color = row[5].trim();
    const range = { label: label, data: [Number(high) - Number(low)], backgroundColor: colorMap[color] };
    ranges.push(range);
  }
  return ranges;
}

export async function getSectionMapping(context) {
  context.document.body.tables.load();
  await context.sync();
  const table = await getTable(context, "Section Name");
  table.load(["values"]);
  await context.sync();

  const elements = table.values;
  const map = new Map();
  for (let i = 0; i < elements.length; i += 1) {
    const row = elements[i];
    const key = row[0];
    if (!key.startsWith("#")) {
      continue;
    }
    const metric = row[1].trim();
    const low = Number(row[2]);
    const high = Number(row[3]);
    let value = [metric, low, high];
    map.set(key.trim(), value);
  }
  // table.delete();
  return map;
}

// Removes section if:
// 1. No section name present in section mapping table
// 2. No metric present in value mapping table
// 3. Metric is outside of [low, high]
export async function checkSection(valueMapping, sectionMapping, p) {
  if (p.tableNestingLevel > 0 || !p.text.startsWith("#") || !p.text.endsWith("#")) {
    return;
  }
  const sectionName = p.text;
  console.log(`Checking section ${sectionName}`);
  if (!sectionMapping.has(sectionName)) {
    console.log(`Section mapping missing: ${sectionName}`);
    await deleteSection(sectionName);
    return;
  }
  const selectionValues = sectionMapping.get(sectionName);
  const metric = selectionValues[0];
  const low = selectionValues[1];
  const high = selectionValues[2];
  if (!valueMapping.has(metric)) {
    console.log(`Value mapping missing: ${metric}`);
    await deleteSection(sectionName);
    return;
  }
  const clientValue = Number(valueMapping.get(metric));
  console.log(`Range check ${clientValue} [${low}, ${high}]`);
  if (clientValue < low || clientValue > high) {
    console.log(`Out of range ${clientValue}`);
    await deleteSection(sectionName);
    return;
  }
}

export async function checkAllSections(valueMapping) {
  return Word.run(async (context) => {
    const sectionMapping = await getSectionMapping(context);
    console.log(sectionMapping);
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    await paragraphs.load();
    await context.sync();
    paragraphs.items.forEach(async (p) => {
      await checkSection(valueMapping, sectionMapping, p);
    });
    await context.sync();
  });
}

export async function cleanUp() {
  return Word.run(async (context) => {
    const sectionMapping = await getSectionMapping(context);
    console.log(sectionMapping);
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    await paragraphs.load();
    await context.sync();
    paragraphs.items.forEach((p) => {
      if (p.text.startsWith("#") && p.text.endsWith("#") && p.tableNestingLevel < 1) {
        p.delete();
      }
    });
    const valueTable = await getTable(context, "Tag");
    valueTable.delete();
    const sectionTable = await getTable(context, "Section Name");
    sectionTable.delete();
    await context.sync();
    await paragraphs.load();
    await context.sync();
    paragraphs.items[0].delete();
    paragraphs.items[1].delete();
    await context.sync();
  });
}

// console.log logs to "Debug Console"
export async function run() {
  return Word.run(async (context) => {
    // Example of API use.
    // const data = await getAPIData();
    // console.log(data);
    // await replacePlaceholder(context, "[Species]", data["species"]["name"]);

    // Extract mapping from tags table.
    const mapping = await getValueMapping(context);
    console.log(mapping);

    // Remove all sections that do not apply.
    await checkAllSections(mapping);

    // Create all charts using values and ranges table.
    for (const tag of Array.from(new Set(mapping.keys()))) {
      const value = mapping.get(tag);
      const ranges = await getRangesForTag(context, tag);
      if (ranges) {
        await createChart(`@${tag.slice(1, -1)}`, value, ranges);
      }
    }

    // Replace all placeholder values with their value.
    for (const tag of Array.from(new Set(mapping.keys()))) {
      await replacePlaceholder(context, tag, mapping.get(tag));
    }

    // Remove input data tables and section name tags.
    await cleanUp();

    await context.sync();
  });
}
