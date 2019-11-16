/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

const DEAD = 0
const LIVE = 1

const COLORS = [
  '#FFFFFF', // dead
  '#000000', // live
]

const propertyToState = props => +(props.format.fill.color === COLORS[LIVE])

const load = async (context, range) => {

  range.load(['address', 'columnCount', 'rowCount'])
  const cellProps = range.getCellProperties({
    format: {
      fill: {
        color: true
      },
    },
  })

  await context.sync()

  const cx = range.columnCount
  const cy = range.rowCount
  const states = cellProps.value.map(row => row.map(propertyToState))

  return {
    context,
    range,
    cx,
    cy,
    states,
  }

}

const stop = () => {

}

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("stop").onclick = stop;
  }
});

export async function run() {
  try {
    await Excel.run(async context => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      const game = await load(context, range)

      console.log(game)
      
    });
  } catch (error) {
    console.error(error);
  }
}
