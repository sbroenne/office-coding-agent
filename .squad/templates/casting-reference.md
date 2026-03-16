# Casting & Persistent Naming Reference

Reference for the coordinator's on-demand casting system. Load this when creating a squad, adding a new AI member, migrating an older team into casting, or explaining how deterministic universe selection works.

## Purpose

Casting gives AI team members memorable, persistent names without changing their actual voice, role, or behavioral contract.

- One fictional universe is selected per assignment.
- Names are persistent identifiers, not role-play instructions.
- Scribe, Ralph, and `@copilot` are exempt from casting.
- The chosen universe is deterministic for the same assignment inputs, except for the LRU tie-break behavior.

This document defines the full v0.8.25 reference model: the complete universe table, scoring algorithm, overflow rules, and the schemas for the state files in `.squad/casting/`.

## Core Rules

1. **One universe per assignment. Never mix universes inside the same cast.**
2. **Names must be unique within the repo.** Retired names remain reserved.
3. **Casting does not change the agent's role.** A Frontend Dev is still a Frontend Dev even if they are named from a heist universe.
4. **Keep the joke quiet.** Never explain the mapping rationale to the user unless they explicitly ask for the casting system.
5. **Prefer names that imply pressure, function, or consequence** rather than literal title words like `Leader` or `Doctor`.

## Full Universe Allowlist

The reference allowlist contains 31 universes. Capacities represent the expected cast depth before overflow handling is required. Repo-local `policy.json` may ship a smaller subset, but the reference standard is the table below.

| Universe | Capacity | Default shape affinity | Typical resonance cues |
|----------|----------|------------------------|------------------------|
| The Usual Suspects | 6 | crime pressure-cooker | deception, tight ensemble, investigation |
| Reservoir Dogs | 8 | criminal crew | tension, aftermath, contained chaos |
| Alien | 8 | survival crew | pressure, technical danger, containment |
| The Thing | 8 | paranoia survival | isolation, uncertainty, trust breakdown |
| The Goonies | 8 | scrappy mixed crew | adventure, improvisation, camaraderie |
| Firefly | 10 | found-family crew | rebels, mission work, odd jobs |
| The Matrix | 10 | rebel cell | systems, reality, operators, control |
| Blade Runner | 10 | noir investigation | ambiguity, urban systems, pursuit |
| Leverage | 10 | specialist crew | role complementarity, fast capers |
| Cowboy Bebop | 10 | drifting crew | bounty work, episodic problem solving |
| Breaking Bad | 12 | escalating duo-plus | consequences, chemistry, moral pressure |
| Star Wars | 12 | rebellion ensemble | factions, mission work, legacy conflict |
| Battlestar Galactica | 12 | fleet under pressure | survival, command, logistics |
| Andor / Rogue One | 12 | insurgency cell | tactical planning, sacrifice, covert work |
| Ocean's Eleven | 14 | heist ensemble | specialization, confidence, coordination |
| Arrested Development | 15 | dysfunctional family/team | comedy, chaos, overlapping agendas |
| Lost | 18 | stranded ensemble | mystery, dependency, rotating focus |
| DC Universe | 18 | superhero network | archetypes, specialties, escalation |
| Parks and Recreation | 18 | civic office ensemble | planning, collaboration, optimism |
| Brooklyn Nine-Nine | 18 | precinct ensemble | procedural work, humor, specialization |
| The Simpsons | 20 | broad town ensemble | recognizability, broad archetypes |
| Lord of the Rings | 20 | fellowship plus worlds | questing, duty, burden, lore |
| Mission: Impossible | 20 | mission team | precise roles, stealth, execution |
| X-Men | 20 | gifted school/team | powers as specialties, mentorship |
| Halo | 20 | military sci-fi roster | operations, chain of command, combat |
| Harry Potter | 22 | school-and-war ensemble | houses, growth, legacy, faction tension |
| Star Trek | 22 | starship/institution | roles, exploration, diplomacy, systems |
| The Expanse | 22 | political-system ensemble | factions, realism, technical stakes |
| Marvel Cinematic Universe | 25 | very broad hero network | huge range, crossovers, flexible depth |
| Game of Thrones | 25 | house-and-war ensemble | politics, strategy, rivalry |
| Mass Effect | 25 | squad-based mission network | specialist companions, choices, systems |

Capacity range across the reference allowlist is **6 to 25**.

## Inputs to Selection

The deterministic selection algorithm consumes four signal groups:

1. **Team size** — how many castable AI members are needed now.
2. **Assignment shape** — the structural pattern of the team (heist crew, office ensemble, rebel cell, survival crew, institution, etc.).
3. **Resonance cues** — words or repo signals from the user's description and current project context.
4. **Recent usage history** — the least-recently-used factor from `history.json`.

### Team size

Count only castable AI members.

- Include Lead, Frontend, Backend, Tester, Product, DevOps, Docs, etc.
- Exclude `Scribe`, `Ralph`, and `@copilot`.
- Exclude humans.

### Assignment shape

Normalize the proposed roster into one dominant shape:

| Shape | Examples |
|-------|----------|
| `crew` | Lead + specialists solving a bounded mission |
| `office` | Product, design, engineering, QA collaboration |
| `rebel-cell` | small async team, infra, platform, ops, systems |
| `survival` | crisis response, stabilization, debugging under pressure |
| `institution` | large structured org, many stable functions |
| `mystery-ensemble` | exploratory or research-heavy assignments |

### Resonance cues

Pull cues from:

- repo name and README themes
- user phrasing (`mission`, `heist`, `dashboard`, `retro`, `rebel`, `incident`)
- architecture terms (`fleet`, `panel`, `ops`, `bridge`, `workspace`)
- squad history if a recurring domain is obvious

Do not overfit. Resonance is a tie-break enhancer, not the main driver.

## Deterministic Selection Algorithm

The scoring model is:

```text
score = size_fit + shape_fit + resonance_fit + lru_bonus
```

Higher score wins.

### 1. Size fit

Choose universes that comfortably hold the team size without needless excess.

Recommended formula:

```text
size_gap = capacity - requested_cast_size
size_fit =
  40  if size_gap == 0
  32  if size_gap == 1
  24  if size_gap == 2
  16  if size_gap between 3 and 5
   8  if size_gap between 6 and 9
   0  if size_gap >= 10
 -40  if capacity < requested_cast_size
```

Interpretation:

- Exact or near-exact fits are strongly preferred.
- Overly large universes are allowed but score lower.
- Universes that are too small are penalized but still remain candidates when overflow handling is acceptable.

### 2. Shape fit

Each universe has one or more default shape affinities.

```text
shape_fit =
  30  strong match
  18  acceptable adjacent match
   0  neutral
 -12  actively awkward fit
```

Examples:

- `Ocean's Eleven` strongly matches `crew`.
- `Parks and Recreation` strongly matches `office`.
- `Alien` strongly matches `survival`.
- `Star Trek` strongly matches `institution`.

### 3. Resonance fit

Award points for keyword or theme alignment.

```text
resonance_fit = min(20, 5 * matched_resonance_signals)
```

Examples:

- Project described as `mission`, `specialists`, `heist`, `pulling jobs` → `Ocean's Eleven`, `Leverage`, `Mission: Impossible`
- Project described as `dashboard`, `ship`, `bridge`, `fleet`, `captain` → `Star Trek`, `Battlestar Galactica`, `The Expanse`
- Project described as `survival`, `incident`, `containment`, `debugging under pressure` → `Alien`, `The Thing`

### 4. LRU bonus

Least-recently-used breaks ties so the same universe does not dominate across assignments.

```text
lru_bonus =
  10  if never used
   6  if not used in last 10 assignments
   3  if not used in last 5 assignments
   0  otherwise
```

When the top score ties, sort by:

1. highest total score
2. highest `size_fit`
3. highest `shape_fit`
4. highest `resonance_fit`
5. highest `lru_bonus`
6. alphabetically by universe name

That final alphabetical tie-break makes the result fully deterministic.

## Reference Pseudocode

```python
def choose_universe(policy, requested_cast_size, shape, resonance_signals, history):
    candidates = []
    for universe in policy["allowlist_universes"]:
        capacity = policy["universe_capacity"][universe]
        score = (
            compute_size_fit(capacity, requested_cast_size)
            + compute_shape_fit(universe, shape)
            + compute_resonance_fit(universe, resonance_signals)
            + compute_lru_bonus(universe, history)
        )
        candidates.append((universe, score))

    return sorted(
        candidates,
        key=lambda item: (
            -item[1],
            -compute_size_fit(policy["universe_capacity"][item[0]], requested_cast_size),
            -compute_shape_fit(item[0], shape),
            -compute_resonance_fit(item[0], resonance_signals),
            -compute_lru_bonus(item[0], history),
            item[0],
        ),
    )[0][0]
```

## Name Allocation Rules

After selecting a universe:

1. Build the available character list for that universe.
2. Remove any names already present in `registry.json`, even if retired.
3. Exclude names reserved by policy or migration rules.
4. Allocate one name per castable agent.
5. Write the mapping to `registry.json`.
6. Write the assignment snapshot to `history.json`.

### Exempt names

- `Scribe` → always `Scribe`
- `Ralph` → always `Ralph`
- `@copilot` → always `@copilot`

## Overflow Handling

If the cast size exceeds the immediately obvious name pool for the chosen universe, do **not** switch universes mid-assignment.

Apply these rules in order:

### 1. Diegetic expansion

Use recurring, minor, or peripheral characters from the same universe.

### 2. Thematic promotion

Expand within the closest natural family of the universe without announcing the promotion.

Examples:

- `Star Wars` → include prequel, sequel, or adjacent rebellion-era characters
- `Marvel Cinematic Universe` → broaden to wider MCU roster
- `Halo` → expand from Spartans to command, ops, and ONI-adjacent characters

### 3. Structural mirroring

Choose names that preserve recognizable role contrast inside the same universe family.

### Overflow principles

- Existing agents are never renamed.
- Overflow is local to the current assignment.
- The universe in `history.json` remains the original selected universe.
- Note overflow in the snapshot metadata when it materially changes name sourcing.

## Casting State Directory

The authoritative state lives in `.squad/casting/`.

| File | Role |
|------|------|
| `policy.json` | authoritative configuration and allowlist |
| `registry.json` | authoritative name registry for members |
| `history.json` | append-only usage history and assignment snapshots |

## `policy.json` Schema

Purpose: define what universes are allowed and how much capacity each has.

```json
{
  "$schema": "https://json-schema.org/draft/2020-12/schema",
  "title": "Squad casting policy",
  "type": "object",
  "required": ["casting_policy_version", "allowlist_universes", "universe_capacity"],
  "properties": {
    "casting_policy_version": { "type": "string" },
    "allowlist_universes": {
      "type": "array",
      "items": { "type": "string" },
      "minItems": 1,
      "uniqueItems": true
    },
    "universe_capacity": {
      "type": "object",
      "additionalProperties": { "type": "integer", "minimum": 6, "maximum": 25 }
    },
    "notes": { "type": "string" }
  },
  "additionalProperties": false
}
```

Minimal example:

```json
{
  "casting_policy_version": "1.1",
  "allowlist_universes": ["The Usual Suspects", "Ocean's Eleven"],
  "universe_capacity": {
    "The Usual Suspects": 6,
    "Ocean's Eleven": 14
  }
}
```

## `registry.json` Schema

Purpose: persist the chosen name for each agent and reserve retired names.

```json
{
  "$schema": "https://json-schema.org/draft/2020-12/schema",
  "title": "Squad casting registry",
  "type": "object",
  "required": ["agents"],
  "properties": {
    "agents": {
      "type": "object",
      "additionalProperties": {
        "type": "object",
        "required": ["persistent_name", "role", "universe", "created_at", "legacy_named", "status"],
        "properties": {
          "persistent_name": { "type": "string" },
          "role": { "type": "string" },
          "universe": { "type": "string" },
          "created_at": { "type": "string", "format": "date-time" },
          "legacy_named": { "type": "boolean" },
          "status": { "type": "string", "enum": ["active", "retired"] },
          "retired_at": { "type": "string", "format": "date-time" },
          "notes": { "type": "string" }
        },
        "additionalProperties": false
      }
    }
  },
  "additionalProperties": false
}
```

Example:

```json
{
  "agents": {
    "Harmony": {
      "persistent_name": "Harmony",
      "role": "Lead",
      "universe": "Ocean's Eleven",
      "created_at": "2026-03-15T09:52:00Z",
      "legacy_named": false,
      "status": "active"
    }
  }
}
```

## `history.json` Schema

Purpose: track universe reuse and preserve assignment snapshots.

```json
{
  "$schema": "https://json-schema.org/draft/2020-12/schema",
  "title": "Squad casting history",
  "type": "object",
  "required": ["universe_usage_history", "assignment_cast_snapshots"],
  "properties": {
    "universe_usage_history": {
      "type": "array",
      "items": {
        "type": "object",
        "required": ["universe", "used_at", "agent_count"],
        "properties": {
          "universe": { "type": "string" },
          "used_at": { "type": "string", "format": "date-time" },
          "agent_count": { "type": "integer", "minimum": 1 },
          "assignment_id": { "type": "string" }
        },
        "additionalProperties": false
      }
    },
    "assignment_cast_snapshots": {
      "type": "object",
      "additionalProperties": {
        "type": "object",
        "required": ["assignment_id", "universe", "created_at", "agents"],
        "properties": {
          "assignment_id": { "type": "string" },
          "universe": { "type": "string" },
          "created_at": { "type": "string", "format": "date-time" },
          "overflow_mode": { "type": "string" },
          "agents": {
            "type": "array",
            "items": {
              "type": "object",
              "required": ["name", "role"],
              "properties": {
                "name": { "type": "string" },
                "role": { "type": "string" }
              },
              "additionalProperties": false
            }
          }
        },
        "additionalProperties": false
      }
    }
  },
  "additionalProperties": false
}
```

Example:

```json
{
  "universe_usage_history": [
    {
      "universe": "Ocean's Eleven",
      "used_at": "2026-03-15T09:52:00Z",
      "agent_count": 6,
      "assignment_id": "assignment-001"
    }
  ],
  "assignment_cast_snapshots": {
    "assignment-001": {
      "assignment_id": "assignment-001",
      "universe": "Ocean's Eleven",
      "created_at": "2026-03-15T09:52:00Z",
      "agents": [
        { "name": "Harmony", "role": "Lead" },
        { "name": "Ellis", "role": "Product Manager" }
      ]
    }
  }
}
```

## Migration for Already-Squadified Repos

When `.squad/team.md` exists but `.squad/casting/` does not:

1. Do not rename the current members.
2. Initialize `policy.json` from defaults.
3. Populate `registry.json` with each existing AI member and set `legacy_named: true`.
4. Create empty or minimal `history.json`.
5. Apply the full algorithm only for members added after migration.

## Review Checklist

Use this checklist when reviewing casting work:

- [ ] One universe selected per assignment
- [ ] Capacity and tie-break logic are deterministic
- [ ] `Scribe`, `Ralph`, and `@copilot` are excluded from cast size
- [ ] Retired names remain reserved in `registry.json`
- [ ] Overflow keeps the existing universe intact
- [ ] `policy.json`, `registry.json`, and `history.json` match the reference schemas
