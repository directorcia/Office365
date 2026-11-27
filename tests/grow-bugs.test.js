const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const html = fs.readFileSync(path.join(__dirname, "..", "docs", "grow-bugs.html"), "utf8");
const scripts = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)];
const simulation = scripts.at(-1)[1];

function createWorld() {
  let nextFrame;
  let clock = 0;
  let positions = [];
  let bugColors = [];
  let foodColors = [];
  const elements = new Map();
  const noOp = () => {};
  const theme = {
    "--cp-success": "#20c060",
    "--cp-warning": "#e0c020",
    "--cp-danger": "#e02020",
    "--cp-link": "#2040e0"
  };

  function element(id) {
    if (elements.has(id)) return elements.get(id);
    const listeners = {};
    const item = {
      id, value: "", checked: false, textContent: "", disabled: false,
      dataset: {}, style: {}, options: [], parentElement: { setAttribute: noOp },
      addEventListener(type, handler) { listeners[type] = handler; },
      dispatch(type, data = {}) { listeners[type]?.({ preventDefault: noOp, ...data }); },
      setAttribute: noOp,
      replaceChildren(...options) { this.options = options; },
      reportValidity() { return true; },
      getBoundingClientRect() { return { left: 0, top: 0, width: 960, height: 560 }; },
      getContext() {
        let shape;
        return {
          clearRect() { if (id === "habitat") { positions = []; bugColors = []; foodColors = []; } },
          translate(x, y) { if (id === "habitat") positions.push({ x, y }); },
          beginPath: noOp, arc(x, y, radius) { shape = radius === 3.5 ? "food" : "other"; },
          fill() {
            if (id === "habitat" && shape === "bug") bugColors.push(this.fillStyle);
            if (id === "habitat" && shape === "food") foodColors.push(this.fillStyle);
          },
          stroke: noOp, moveTo: noOp, lineTo: noOp, fillRect: noOp, setLineDash: noOp, closePath: noOp,
          ellipse() { shape = "bug"; }, rotate: noOp, save: noOp, restore: noOp
        };
      }
    };
    elements.set(id, item);
    return item;
  }

  element("bug-lifetime").value = "15";
  element("sim-speed").value = "1";
  element("food-mode").value = "renewable";
  element("renewal-amount").value = "15";
  element("renewal-interval").value = "15";
  element("terrain-mode").value = "open";
  element("predator-count").value = "0";
  element("click-mode").value = "inspect";
  const document = { documentElement: {}, getElementById: element };
  vm.runInNewContext(simulation, {
    document, getComputedStyle: () => ({ getPropertyValue: (name) => theme[name] || "#000000" }),
    requestAnimationFrame: (callback) => { nextFrame = callback; },
    Option: function (text, value) { this.text = text; this.value = value; }
  });

  function advance(count) {
    for (let i = 0; i < count; i++) {
      clock += 50;
      nextFrame(clock);
    }
  }
  function start({ bugs = 2, food = 0, lifetime = 5, seed = "lifetime-test", evolution = false, renewable = false, renewalAmount = 15, renewalInterval = 15, terrain = "open", predators = 0 } = {}) {
    element("start-bugs").value = String(bugs);
    element("start-food").value = String(food);
    element("bug-lifetime").value = String(lifetime);
    element("seed").value = seed;
    element("evolution").checked = evolution;
    element("food-mode").value = renewable ? "renewable" : "finite";
    element("renewal-amount").value = String(renewalAmount);
    element("renewal-interval").value = String(renewalInterval);
    element("terrain-mode").value = terrain;
    element("predator-count").value = String(predators);
    element("settings").dispatch("submit");
  }
  function inspect(id) {
    element("bug-select").value = String(id);
    element("bug-select").dispatch("change");
    return element("inspector").textContent;
  }
  function feed(index) {
    element("click-mode").value = "food";
    const position = positions[index];
    element("habitat").dispatch("click", { clientX: position.x, clientY: position.y });
  }
  return { element, advance, start, inspect, feed, positions: () => positions, bugColors: () => bugColors, foodColors: () => foodColors };
}

test("a contact produces one nearby child without repeated births during overlap", () => {
  const world = createWorld();
  world.start({ bugs: 2, lifetime: 120, seed: "late-6" });
  world.feed(0);
  world.advance(1);
  world.feed(1);
  world.advance(1);
  world.advance(61);
  assert.equal(world.element("bugs").textContent, 3);
  assert.equal(world.element("contact-births").textContent, 1);
  const [a, b, child] = world.positions();
  assert.ok(Math.hypot(a.x - b.x, a.y - b.y) <= 16);
  assert.ok(Math.hypot(a.x - child.x, a.y - child.y) <= 16);
  world.advance(3);
  assert.equal(world.element("bugs").textContent, 3);
  assert.equal(world.element("contact-births").textContent, 1);
});

test("collisions involving hungry or near-starving bugs do not produce children", () => {
  for (const [seed, frames] of [["late-6", 63], ["late-436", 89]]) {
    const world = createWorld();
    world.start({ bugs: 2, lifetime: 5, seed });
    world.advance(frames);
    const [a, b] = world.positions();
    assert.ok(Math.hypot(a.x - b.x, a.y - b.y) <= 16);
    assert.notEqual(world.bugColors()[0], "rgb(32, 192, 96)");
    assert.notEqual(world.bugColors()[1], "rgb(32, 192, 96)");
    assert.equal(world.element("bugs").textContent, 2);
    assert.equal(world.element("contact-births").textContent, 0);
  }
});

test("a green bug colliding with a hungry bug does not produce a child", () => {
  const world = createWorld();
  world.start({ bugs: 2, lifetime: 5, seed: "late-6" });
  world.advance(51);
  world.feed(0);
  world.advance(12);
  const [a, b] = world.positions();
  assert.ok(Math.hypot(a.x - b.x, a.y - b.y) <= 17);
  assert.equal(world.bugColors()[0], "rgb(32, 192, 96)");
  assert.notEqual(world.bugColors()[1], "rgb(32, 192, 96)");
  assert.equal(world.element("bugs").textContent, 2);
  assert.equal(world.element("contact-births").textContent, 0);
});

test("contact births respect the population cap and inherit traits when evolution is enabled", () => {
  const world = createWorld();
  world.start({ bugs: 100, lifetime: 120, seed: "pair-1875" });
  world.advance(2);
  assert.equal(world.element("bugs").textContent, 100);
  assert.equal(world.element("contact-births").textContent, 0);
  world.start({ bugs: 2, lifetime: 120, seed: "late-6", evolution: true });
  world.feed(0);
  world.advance(1);
  world.feed(1);
  world.advance(62);
  assert.equal(world.element("contact-births").textContent, 1);
  assert.match(world.inspect(3), /Generation 2/);
});

test("offspring inherit the exact genes of their chosen parent when mutation is disabled", () => {
  const world = createWorld();
  world.start({ bugs: 2, lifetime: 120, seed: "late-6", evolution: false });
  const parents = {
    1: world.inspect(1).match(/Speed gene ([\d.]+).*Turn gene ([\d.]+)/),
    2: world.inspect(2).match(/Speed gene ([\d.]+).*Turn gene ([\d.]+)/)
  };
  world.feed(0);
  world.advance(1);
  world.feed(1);
  world.advance(62);
  const childText = world.inspect(3);
  const child = childText.match(/Parent #(\d+).*Speed gene ([\d.]+).*Turn gene ([\d.]+)/);
  assert.ok(child);
  assert.equal(child[2], parents[child[1]][1]);
  assert.equal(child[3], parents[child[1]][2]);
});

test("terrain renders reproducibly and predators can kill and trigger replacement", () => {
  const world = createWorld();
  world.start({ bugs: 3, food: 300, lifetime: 120, seed: "terrain-seed", terrain: "mixed", predators: 0 });
  world.advance(20);
  assert.match(world.element("habitat-description").textContent, /with terrain$/);
  assert.equal(world.element("predator-count").value, "0");
  world.start({ bugs: 3, food: 300, lifetime: 120, seed: "hunt-1", terrain: "open", predators: 1 });
  world.advance(300);
  assert.match(world.element("habitat-description").textContent, /and 1 predator/);
  assert.ok(Number(world.element("predation-deaths").textContent) > 0);
  assert.match(world.element("population-summary").textContent, /hunted/);
  assert.ok(Number(world.element("bugs").textContent) > 0);
});

test("shelters protect a bug even when a predator overlaps it", () => {
  const world = createWorld();
  world.start({ bugs: 1, food: 300, lifetime: 120, seed: "shelter-15108", terrain: "mixed", predators: 1 });
  const [bug, predator] = world.positions();
  assert.ok(Math.hypot(bug.x - predator.x, bug.y - predator.y) < 20);
  assert.ok(Math.hypot(bug.x - 830, bug.y - 455) < 45);
  world.advance(1);
  assert.equal(world.element("predation-deaths").textContent, 0);
  assert.equal(world.element("bugs").textContent, 1);
});

test("population, generation, and selection charts record samples", () => {
  const world = createWorld();
  world.start({ bugs: 20, food: 100, lifetime: 20, seed: "selection", evolution: true });
  world.advance(301);
  assert.match(world.element("population-summary").textContent, /alive.*births.*starved.*hunted/);
  assert.match(world.element("generation-summary").textContent, /gen/);
  assert.match(world.element("selection-summary").textContent, /speed .*% · turning .*%/);
});

test("renewable food appears at the configured interval, while finite food does not", () => {
  const world = createWorld();
  const settings = { bugs: 1, lifetime: 120, renewable: true, renewalAmount: 7, renewalInterval: 2 };
  world.start(settings);
  world.advance(40);
  assert.equal(world.element("food").textContent, 0);
  world.advance(1);
  assert.equal(world.element("food").textContent, 7);
  world.start({ ...settings, renewable: false });
  world.advance(41);
  assert.equal(world.element("food").textContent, 0);
});

test("sprinkling works while paused and respects the food limit", () => {
  const world = createWorld();
  world.start({ food: 298, lifetime: 120 });
  world.element("pause").dispatch("click");
  world.element("sprinkle").dispatch("click");
  assert.equal(world.element("food").textContent, 300);
  assert.match(world.element("sprinkle-status").textContent, /Added 2 food pieces/);
  world.element("sprinkle").dispatch("click");
  assert.equal(world.element("food").textContent, 300);
  assert.match(world.element("sprinkle-status").textContent, /already has 300/);
});

test("food is blue and bug colors fade from green through amber toward red, then reset on feeding", () => {
  const world = createWorld();
  world.start({ bugs: 1, lifetime: 8 });
  assert.equal(world.bugColors()[0], "rgb(32, 192, 96)");
  world.advance(41);
  assert.equal(world.bugColors()[0], "rgb(32, 192, 96)");
  world.advance(20);
  assert.equal(world.bugColors()[0], "rgb(128, 192, 64)");
  world.advance(20);
  assert.equal(world.bugColors()[0], "rgb(224, 192, 32)");
  world.advance(79);
  assert.equal(world.bugColors()[0], "rgb(224, 32, 32)");
  world.feed(0);
  assert.equal(world.foodColors()[0], "#2040e0");
  world.advance(1);
  assert.equal(world.bugColors()[0], "rgb(32, 192, 96)");
});

test("eating twice does not create a bug without a green collision", () => {
  const world = createWorld();
  world.start({ bugs: 1, lifetime: 120, evolution: true });
  world.feed(0);
  world.advance(2);
  world.feed(0);
  world.advance(1);
  assert.match(world.inspect(1), /Meals 2/);
  assert.equal(world.element("bugs").textContent, 1);
  assert.equal(world.element("contact-births").textContent, 0);
});

test("green contact without stored meals cannot reproduce", () => {
  const world = createWorld();
  world.start({ bugs: 2, lifetime: 120, seed: "pair-1875" });
  world.advance(5);
  assert.equal(world.element("bugs").textContent, 2);
  assert.equal(world.element("contact-births").textContent, 0);
});

test("a birth spends both parents' stored meals", () => {
  const world = createWorld();
  world.start({ bugs: 2, lifetime: 120, seed: "late-6", evolution: true });
  world.feed(0);
  world.advance(1);
  world.feed(1);
  world.advance(62);
  assert.match(world.inspect(1), /Birth energy 0/);
  assert.match(world.inspect(2), /Birth energy 0/);
  assert.equal(world.element("contact-births").textContent, 1);
});

test("food adds health to the eater while another bug's health continues to drain", () => {
  const world = createWorld();
  world.start();
  world.advance(61);
  assert.match(world.inspect(1), /Health 2\.0/);
  world.feed(0);
  world.advance(1);
  assert.match(world.inspect(1), /Meals 1 · Birth energy 1 · Health 7\.0/);
  assert.match(world.inspect(2), /Meals 0 · Birth energy 0 · Health 2\.0/);
  world.advance(40);
  assert.deepEqual(world.element("bug-select").options.map((option) => option.value), ["", "1", "3"]);
  assert.equal(world.element("replaced").textContent, 1);
  assert.match(world.element("survival-summary").textContent, /50\.0%/);
  assert.equal(world.element("status-text").textContent, "Simulation running");
});

test("multiple meals stack health instead of resetting it to one portion", () => {
  const world = createWorld();
  world.start({ bugs: 1, lifetime: 5 });
  world.feed(0);
  world.advance(1);
  world.feed(0);
  world.advance(1);
  assert.match(world.inspect(1), /Meals 2 · Birth energy 2 · Health 14\.9/);
});

test("all bugs starving at once ends the simulation without replacements", () => {
  const world = createWorld();
  world.start();
  world.advance(101);
  assert.equal(world.element("elapsed").textContent, "5.0");
  assert.equal(world.element("bugs").textContent, 0);
  assert.equal(world.element("replaced").textContent, 0);
  assert.equal(world.element("status-text").textContent, "No bugs survived");
  assert.match(world.element("survival-summary").textContent, /0\.0%/);
});

test("seeded lifetimes and sampled statistics replay across simulation speeds", () => {
  const world = createWorld();
  world.start({ bugs: 20, food: 100, lifetime: 15, seed: "repeat" });
  world.advance(401);
  const snapshot = ["elapsed", "bugs", "food", "replaced", "contact-births", "survival-summary", "speed-summary", "food-summary"]
    .map((id) => world.element(id).textContent);
  world.element("sim-speed").value = "4";
  world.start({ bugs: 20, food: 100, lifetime: 15, seed: "repeat" });
  world.advance(101);
  assert.deepEqual(
    ["elapsed", "bugs", "food", "replaced", "contact-births", "survival-summary", "speed-summary", "food-summary"]
      .map((id) => world.element(id).textContent),
    snapshot
  );
});
