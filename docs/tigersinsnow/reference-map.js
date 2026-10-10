(function (root) {
  'use strict';
  const COLS = 19, ROWS = 11;
  const terrainNames = { '.': 'plain', F: 'forest', R: 'ridge', T: 'town', W: 'westwall' };
  const terrainRows = [
    '. R . . . . F R R T T R T R R F W F W',
    'R . . F R . . T F F F R R R F F T W W',
    '. . T F F F . F F F F F F F R F F F W',
    '. R F F R R . F F F F F F F F R F . W',
    'R F T R R F F T F F F F F F T R R W F',
    'F F F F F F T F F T F F T R R R R F W',
    'F F F . F F F F F F R R R T R F W W F',
    'F F F F F F F F T R R T F F F R F F W',
    'F F F F R F F F R R R R F F T R W W W',
    'F T F F F R F F T . F R F R R . . . T',
    'F R R T F F F F F F F T R R R R R R R'
  ].map(row => row.split(' '));
  const position = (col, row) => ({ x: col * 2 + 1 - row % 2, y: row });
  const key = tile => `${tile.x},${tile.y}`;
  const edgeKey = (a, b) => [key(a), key(b)].sort().join('|');
  const entries = [
    [1, 2, 0, 'north'], [2, 3, 0, 'north'], [3, 4, 0, 'north'],
    [4, 6, 0, 'north'], [5, 7, 0, 'north'], [6, 8, 0, 'north'],
    [7, 9, 0, 'north'], [8, 11, 0, 'north'], [9, 12, 0, 'north'], [10, 13, 0, 'north'],
    [11, 18, 2, 'german'], [12, 18, 3, 'german'], [13, 18, 4, 'german'],
    [14, 18, 5, 'german'], [15, 18, 6, 'german'], [16, 18, 7, 'german'],
    [17, 15, 10, 'south'], [18, 14, 10, 'south'], [19, 12, 10, 'south'],
    [20, 11, 10, 'south'], [21, 10, 10, 'south'], [22, 8, 10, 'south'],
    [23, 7, 10, 'south'], [24, 6, 10, 'south'],
    [25, 0, 7, 'west'], [26, 0, 4, 'west'], [27, 0, 2, 'west'], [28, 2, 4, 'west']
  ].map(([number, col, row, edge]) => ({ number, col, row, edge, ...position(col, row) }));
  // Shared hex boundaries transcribed from the blue lines on the supplied map display.
  const riverEdges = [
    [6,0,7,1],[7,0,8,1],[7,0,7,1],[8,0,9,1],[8,0,8,1],[9,0,10,0],[9,0,10,1],[9,0,9,1],
    [3,1,3,2],[4,1,4,2],[4,1,3,2],[5,1,5,2],[5,1,4,2],[6,1,7,1],[6,1,6,2],[6,1,5,2],
    [9,1,10,1],[9,1,9,2],[10,1,10,2],[10,1,9,2],[11,1,10,2],
    [0,2,1,3],[0,2,0,3],[1,2,2,3],[1,2,1,3],[2,2,3,2],[2,2,3,3],[2,2,2,3],
    [8,2,9,2],[8,2,9,3],[10,2,11,2],[11,2,12,3],[11,2,11,3],[12,2,13,3],[12,2,12,3],
    [13,2,14,3],[13,2,13,3],[14,2,15,3],[14,2,14,3],
    [2,3,3,3],[2,3,2,4],[8,3,9,3],[9,3,8,4],[12,3,13,3],[13,3,12,4],[16,3,16,4],
    [1,4,2,4],[2,4,2,5],[8,4,9,4],[9,4,9,5],[12,4,13,4],[15,4,16,4],[15,4,16,5],
    [2,5,3,5],[2,5,2,6],[3,5,3,6],[3,5,2,6],[4,5,4,6],[4,5,3,6],[5,5,5,6],[5,5,4,6],
    [6,5,6,6],[6,5,5,6],[7,5,7,6],[7,5,6,6],[9,5,10,5],[10,5,10,6],[10,5,9,6],
    [11,5,11,6],[11,5,10,6],[12,5,11,6],[14,5,13,6],[15,5,16,5],[15,5,15,6],
    [1,6,2,6],[1,6,2,7],[4,6,5,6],[5,6,5,7],[9,6,10,6],[9,6,10,7],
    [13,6,14,6],[13,6,14,7],[14,6,15,6],[15,6,15,7],
    [1,7,2,7],[1,7,1,8],[2,7,1,8],[5,7,6,7],[9,7,10,7],[9,7,9,8],
    [13,7,14,7],[14,7,13,8],[15,7,16,7],[16,7,15,8],[17,7,18,7],[18,7,17,8],
    [0,8,1,8],[1,8,2,8],[1,8,1,9],[2,8,3,9],[2,8,2,9],[3,8,4,9],[3,8,3,9],[4,8,4,9],
    [8,8,9,8],[11,8,12,9],[12,8,13,9],[12,8,12,9],[13,8,14,8],[13,8,14,9],[13,8,13,9],
    [14,8,15,9],[14,8,14,9],[15,8,16,8],[15,8,16,9],[15,8,15,9],
    [16,8,17,9],[16,8,16,9],[17,8,18,8],[17,8,18,9],[17,8,17,9],[18,8,18,9],
    [1,9,2,9],[2,9,2,10],[2,9,1,10],[3,9,3,10],[3,9,2,10],[4,9,5,9],[4,9,4,10],
    [5,9,5,10],[5,9,4,10],[6,9,6,10],[6,9,5,10],[7,9,7,10],[7,9,6,10],
    [8,9,8,10],[8,9,7,10],[9,9,9,10],[9,9,8,10],[10,9,10,10],[10,9,9,10],[11,9,11,10],[11,9,10,10],
    [12,9,12,10],[12,9,11,10],[13,9,14,9],[13,9,13,10],[13,9,12,10],[14,9,14,10],[14,9,13,10],
    [15,9,15,10],[15,9,14,10],[16,9,16,10],[16,9,15,10],[17,9,16,10]
  ].map(([ac, ar, bc, br]) => ({ a: position(ac, ar), b: position(bc, br) }));
  const rivers = new Set(riverEdges.map(edge => edgeKey(edge.a, edge.b)));
  const center = tile => ({ x: 38 + tile.x * Math.sqrt(3) * 13, y: 40 + tile.y * 39 });
  const riverSegments = riverEdges.map(({ a, b }) => {
    const first = center(a), second = center(b), dx = second.x - first.x, dy = second.y - first.y;
    const length = Math.hypot(dx, dy), mx = (first.x + second.x) / 2, my = (first.y + second.y) / 2;
    return [[mx - dy / length * 13, my + dx / length * 13],
      [mx + dy / length * 13, my - dx / length * 13]];
  });
  function createMap() {
    const map = {};
    terrainRows.forEach((row, r) => row.forEach((symbol, c) => {
      const tile = { ...position(c, r), col: c, row: r, terrain: terrainNames[symbol], owner: null };
      if (r === 0 && c <= 12) tile.alternate = 'north';
      if (c === 0 && r > 0) tile.alternate = 'west';
      if (r === 10 && c >= 2 && c <= 16) tile.alternate = 'south';
      if (c === 18 && r < 9) tile.alternate = 'german';
      tile.exit = (r === 0 && c <= 9) || (c === 0 && r <= 7);
      map[key(tile)] = tile;
    }));
    for (const entry of entries) map[key(entry)].entry = entry.number;
    for (const [c, r, name, value, sourcePoints] of [
      [7, 4, 'MARCHE', 2, 25], [6, 5, 'ROCHEFORT', 2, 25], [11, 7, 'BASTOGNE', 3, 50]
    ]) Object.assign(map[key(position(c, r))], { name, value, sourcePoints, owner: 'american' });
    return map;
  }
  const api = { COLS, ROWS, terrainRows, position, entries, riverEdges, riverSegments, center, createMap,
    crossesRiver: (a, b) => rivers.has(edgeKey(a, b)), edgeKey };
  if (typeof module === 'object') module.exports = api;
  else root.ReferenceMap = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
