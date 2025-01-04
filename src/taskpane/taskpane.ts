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

// Finds and replaces placeholder text with value anywhere in the document.
export async function replacePlaceholder(context, placeholder, value) {
  const body = context.document.body;
  const paragraphs = body.paragraphs;
  await paragraphs.load();
  await context.sync();
  paragraphs.items.forEach((p) => {
    const originalText = p.text;
    const newText = originalText.replace(placeholder, value);
    p.clear();
    p.insertText(newText, Word.InsertLocation.start);
  });
  await context.sync();
}

// Extracts a map of key, value pairs from a 2 column table.
export async function getValueMapping(context) {
  context.document.body.tables.load();
  await context.sync();
  const table = context.document.body.tables.getFirst();
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
  return map;
}

// Example of how we can fetch API data.
export async function getAPIData() {
  // Define the API URL
  const apiUrl = "https://pokeapi.co/api/v2/pokemon/ditto";
  const response = await fetch(apiUrl);
  return await response.json();
}

// Inserts an image into the last paragraph.
export async function insertImage(context, base64Image) {
  await context.document.body.paragraphs.load();
  await context.sync();
  context.document.body.paragraphs
    .getLast()
    .insertParagraph("", "After")
    .insertInlinePictureFromBase64(base64Image.replace("data:image/png;base64,", ""), "End");

  await context.sync();
}

// Creates and adds a chart to the document.
export async function createChart(context) {
  console.log("Creating chart");
  // Get the div element
  const div = document.getElementById("renderDiv");

  // Create a new canvas element
  const canvas = document.createElement("canvas");

  // Set any attributes for the canvas (optional)
  // canvas.width = 500;
  // canvas.height = 300;

  // Append the canvas to the div
  div.appendChild(canvas);

  const optimalRange = [20, 50];
  const clientValue = 40;
  const metricName = "Ferratin";

  // Min? max? title?
  const myChart = new Chart(canvas, {
    data: {
      datasets: [
        {
          type: "bar",
          label: "Optimal Range",
          data: [optimalRange],
          indexAxis: "y",
        },
        {
          type: "bar",
          label: "Too High",
          data: [[10, 30]],
          indexAxis: "y",
        },
        {
          type: "line",
          label: "Your Value",
          data: [clientValue],
          indexAxis: "y",
        },
      ],
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
            insertImage(context, myChart.toBase64Image());
            div.removeChild(canvas);
          }
        },
      },
    },
  });
}

// Deletes a section.
export async function deleteSection(context, sectionName) {
  const searchQuery = `[#]${sectionName}*[#]${sectionName}`;

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
  await context.sync();
}

// console.log logs to "Debug Console"
export async function run() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello World!", Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    // Replace all placeholder values with their value.
    // const mapping = await getValueMapping(context);
    // for (const key of Array.from(new Set(mapping.keys()))) {
    //   await replacePlaceholder(context, key, mapping.get(key));
    // }

    // Example of API use.
    // const data = await getAPIData();
    // console.log(data);
    // await replacePlaceholder(context, "[Species]", data["species"]["name"]);

    // createChart(context);

    deleteSection(context, "FerratinLow");

    await context.sync();
  });
}
