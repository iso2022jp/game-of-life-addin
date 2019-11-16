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

const NEIGHBORS = [
  [-1, -1],
  [0, -1],
  [1, -1],
  [-1, 0],
  [1, 0],
  [-1, 1],
  [0, 1],
  [1, 1],
]

const RULE = [
  [ 0, 0, 0, 1, 0, 0, 0, 0, 0, ], // dead
  [ 0, 0, 1, 1, 0, 0, 0, 0, 0, ], // live
]

const range = (count, mapping) => Array.from({length: count}, (_, i) => mapping(i))

const propertyToState = props => +(props.format.fill.color === COLORS[LIVE])
const stateToProperty = state => ({ format: { fill: {color: COLORS[state] }} })

const advance = (state, lives) => RULE[state][lives]

const countNeighbourSurvivors = (states, x, y) => {
  return NEIGHBORS.reduce((count, [dx, dy]) => {
    return count + ((states[y + dy] || [])[x + dx] || DEAD)
  }, 0)
}

const genarate = (states, cx, cy) => {
  return range(cy, y => range(cx, x => {
    const lives = countNeighbourSurvivors(states, x, y)
    return advance(states[y][x], lives)
  }))
}

let currentGame = null
let currentTimer = null

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

const update = async (context, range, states, prevStates) => {
  
  const cellProps = states.map((row, y) => row.map((state, x) => {
    const prevState = prevStates[y][x]
    return state === prevState ? {} : stateToProperty(state)
  }))

  range.setCellProperties(cellProps)  
  await context.sync()
}

const step = () => {
  console.log('Step')

  const states = genarate(currentGame.states, currentGame.cx, currentGame.cy)
  console.log(states)

  const game = {
    ...currentGame,
    states,
  }

  update(game.context, game.range, states, currentGame.states)

  currentGame = game

}

const stop = () => {
  currentTimer && clearInterval(currentTimer)
  currentTimer = null
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

      currentGame = game
      
      console.log('Loaded')
      console.log(game)

      context.trackedObjects.add(range)
      context.sync()

      currentTimer = setInterval(step, 1000)
      
    });
  } catch (error) {
    console.error(error);
  }
}
