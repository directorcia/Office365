(function (root) {
  'use strict';
  const stroke = 'fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"';
  const paths = {
    swords: `<path ${stroke} d="M5 19 L14 6"/><path ${stroke} d="M14 6 l2-2 4 0-2 4-2 2"/><path ${stroke} d="M19 19 L10 6"/><path ${stroke} d="M10 6 L8 4 H4 l2 4 2 2"/>`,
    track: `<path ${stroke} d="M4 16c2-6 4-8 8-8s6 2 8 8"/><path ${stroke} d="M8 16v4M12 16v4M16 16v4"/><path ${stroke} d="M15 8l3-3M18 5h-4"/>`,
    crate: `<rect ${stroke} x="4" y="7" width="16" height="12" rx="1"/><path ${stroke} d="M4 12h16M12 7v12"/>`,
    gauge: `<path ${stroke} d="M4 16a8 8 0 0 1 16 0"/><path ${stroke} d="M12 16l5-6"/><circle fill="currentColor" cx="12" cy="16" r="1.4"/>`,
    cannon: `<circle ${stroke} cx="7" cy="17" r="3"/><path ${stroke} d="M9 15 L19 8M9 17.5 L17 12"/><path ${stroke} d="M16 7l3 1"/>`,
    fuel: `<rect ${stroke} x="6" y="6" width="12" height="14" rx="1.5"/><path ${stroke} d="M9 6V4h6v2M8 12h8"/>`,
    star: `<polygon ${stroke} points="12,3 14.6,8.8 21,9.4 16.2,13.6 17.6,20 12,16.6 6.4,20 7.8,13.6 3,9.4 9.4,8.8"/>`,
    flag: `<path ${stroke} d="M6 3v18"/><path ${stroke} d="M6 4h11l-2.5 4L17 12H6"/>`,
    pin: `<path ${stroke} d="M12 21s7-6.2 7-11a7 7 0 1 0-14 0c0 4.8 7 11 7 11z"/><circle ${stroke} cx="12" cy="10" r="2"/>`,
    armour: `<ellipse ${stroke} cx="12" cy="12" rx="9" ry="6"/>`,
    infantry: `<rect ${stroke} x="4" y="6" width="16" height="12"/><path ${stroke} d="M4 6l16 12M20 6L4 18"/>`,
    airborne: `<path ${stroke} d="M3 9c3-4 6-5 9-5s6 1 9 5"/><rect ${stroke} x="6" y="11" width="12" height="8"/><path ${stroke} d="M6 11l12 8M18 11L6 19"/>`,
    mech: `<ellipse ${stroke} cx="12" cy="12" rx="9" ry="6"/><path ${stroke} d="M6 9l12 6M18 9L6 15"/>`
  };
  function icon(name) {
    return `<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">${paths[name] || paths.star}</svg>`;
  }
  const api = { icon };
  if (typeof module === 'object') module.exports = api;
  else root.PanelIcons = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
