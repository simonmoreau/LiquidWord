/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
import { Liquid } from "liquidjs";

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {

    let ranges: Word.Range[] = [];

    // Search for loops
    const loopSearchString = "([{]% for)*([{]% endfor %[}])";
    await GatherRanges(loopSearchString, context, ranges);

    // Search for ifs
    const ifsSearchString = "([{]% if)*([{]% endif %[}])";
    await GatherRanges(ifsSearchString, context, ranges);

    // Search for unless
    const unlessSearchString = "([{]% unless)*([{]% endunless %[}])";
    await GatherRanges(unlessSearchString, context, ranges);

    // Search for unless
    const valueSearchString = "([{][{])*([}][}])";
    await GatherRanges(valueSearchString, context, ranges);

    // Data
    var data = { name: "simon", title: "Awesome Shoes" };

    await ReplaceText(context, ranges, data);

    await context.sync();

  });
}

async function GatherRanges(searchString: string, context: Word.RequestContext, ranges: Word.Range[]) {
  let foundItems: Word.RangeCollection = context.document.body.search(searchString, { matchWildcards: true }).load();
  await context.sync();

  for (let i = 0; i < foundItems.items.length; i++) {
    let range: Word.Range = foundItems.items[i];
    // Add the range to the list
    ranges.push(range);
  }
}

async function ReplaceText(context: Word.RequestContext, ranges: Word.Range[], data: any) {
  // Create the parsing engine
  const engine: Liquid = new Liquid();

  for (let i = 0; i < ranges.length; i++) {
    let range: Word.Range = ranges[i];
    // Queue a command to get the HTML contents of the paragraph.
    var html = range.getHtml();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    try {
      await context.sync();
      const htmlText = html.value.replace('&quot;','"').replace('&quot;','"'); 

      let value: any = await engine.parseAndRender(htmlText, data);
      range.insertHtml(value, Word.InsertLocation.replace);

      await context.sync();
    }
    catch(e) {
      console.log(e);
    }
  }
}

function FilterRanges(ranges: Word.Range[]){

  for (let i = 0; i < ranges.length; i++) {
    let range: Word.Range = ranges[i];
    
    
  }
}
