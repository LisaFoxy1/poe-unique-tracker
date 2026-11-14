import {
  useEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
  type KeyboardEvent as ReactKeyboardEvent,
} from "react";
import { open } from "@tauri-apps/plugin-dialog";
import { fetch } from "@tauri-apps/plugin-http";
import { readFile } from "@tauri-apps/plugin-fs";
import { appConfigDir } from "@tauri-apps/api/path";
import Database from "@tauri-apps/plugin-sql";
import { emit, listen } from "@tauri-apps/api/event";
import { WebviewWindow } from "@tauri-apps/api/webviewWindow";
import {
  cursorPosition,
  LogicalPosition,
  monitorFromPoint,
  primaryMonitor,
} from "@tauri-apps/api/window";
import { invoke } from "@tauri-apps/api/core";
import { readText, writeText } from "@tauri-apps/plugin-clipboard-manager";
import { register, unregister } from "@tauri-apps/plugin-global-shortcut";
import { read, utils } from "xlsx";
import {
  ensureEditionAvailabilitySchema,
  isCanonicalCatalogueReady,
  reconcileImportedCollection,
  shouldCheckCatalogueNow,
  syncCatalogueFromPoeWiki,
  type EditionAvailability,
  type ImportReconciliationSummary,
} from "./catalogueSync";
import {
  identifyUniqueFromClipboard,
  type ItemIdentificationResult,
} from "./poeItemParser";
import ImportReviewModal from "./ImportReviewModal";
import AppUpdater from "./AppUpdater";
import { getVersion } from "@tauri-apps/api/app";
import {
  chooseTrackerSnapshot,
  createTrackerBackup,
  exportTrackerSnapshot,
  restoreTrackerSnapshot,
  type TrackerSnapshotPreview,
} from "./trackerSnapshot";
import "./App.css";

type TrackingFlag =
  | "owned"
  | "wearing"
  | "foil"
  | "foulborn"
  | "vestigial";

type ExtraVariant = "foil" | "foulborn" | "vestigial";

type EditionAvailabilityMap = Record<
  ExtraVariant,
  EditionAvailability
>;

type EditionSourceMap = Record<
  ExtraVariant,
  string
>;

type UniqueEntry = {
  id: string;
  familyId: string;
  name: string;
  baseType: string | null;
  itemType: string;
  variantLabel: string | null;
  importStatus: string;
  flags: TrackingFlag[];
  editionAvailability: EditionAvailabilityMap;
  editionSources: EditionSourceMap;
  reviewed: boolean;
  isLegacyOnly: boolean;
  catalogueSource: string;
};

type CatalogueUpdateSummary = {
  revision: string;
  label: string;
  newFamilies: number;
  newVariants: number;
  dropDisabled: number;
  updatedEntries: number;
};

type SortMode = "alphabetical" | "type";

type StatusFilter =
  | "all"
  | "unreviewed"
  | "missing"
  | TrackingFlag;

type StatusKey = "missing" | TrackingFlag;

type StatusColors = Record<StatusKey, string>;

type CollectionRules = Record<TrackingFlag, boolean>;

type ExtraTracking = Record<ExtraVariant, boolean>;

type CollectionProfile = {
  id: string;
  name: string;
  kind: "standard" | "challenge";
  isArchived: boolean;
};




type PoeTradeLeagueEntry = {
  id: string;
  startAt: string | null;
  endAt: string | null;
  event?: boolean;
  realm?: string | null;
  category?: {
    id: string;
    current?: boolean;
  };
  rules?: {
    id: string;
    name?: string;
    description?: string;
  }[];
};



type PoeLeagueDiscovery = {
  activeLeagues: PoeTradeLeagueEntry[];
  endedLeagueKeys: Set<string>;
};


type WikiPageTitleResponse = {
  query?: {
    pages?: Array<{
      pageid?: number;
      title?: string;
    }>;
  };
};

type LeagueRolloverChange = {
  uniqueId: string;
  name: string;
  addedFlags: TrackingFlag[];
};

type LeagueRolloverPreview = {
  profileId: string;
  oldLeagueName: string;
  newLeagueName: string;
  changedUniques: number;
  flagCounts: Record<TrackingFlag, number>;
  changes: LeagueRolloverChange[];
};

type LeagueRolloverMode =
  | "dev-pending"
  | "dev-complete"
  | "pending"
  | "complete";

type SqliteTransactionStatement = {
  sql: string;
  params?: string[];
};

type BackupMode = "automatic" | "ask" | "off";

type ImportMode = "status-list" | "missing-only";

type ParsedImportRow = {
  id: string;
  name: string;
  itemType: string;
  rawStatus: string;
};

type ProtectedLeagueUnique = {
  id: string;
  name: string;
  itemType: string;
  variantLabel: string | null;
  releaseVersion: string | null;
};

type PendingImport = {
  fileName: string;
  rows: ParsedImportRow[];
  suggestedMode: ImportMode;
  latestReleaseLine: string | null;
  protectedLeagueUniques: ProtectedLeagueUnique[];
};

type MissingOnlyInferenceSummary = {
  inferredOwned: number;
  explicitMissing: number;
  protectedNewLeague: ProtectedLeagueUnique[];
  protectedAmbiguous: number;
  latestReleaseLine: string | null;
  destinationProfileId: string;
  destinationProfileName: string;
};

const STANDARD_PROFILE_ID = "standard";
const CURRENT_LEAGUE_PROFILE_ID = "current-league";

const IMPORT_SHEET_TYPES: Record<string, string> = {
  amulet: "Amulet",
  amulets: "Amulet",
  axe: "Axe",
  axes: "Axe",
  belt: "Belt",
  belts: "Belt",
  "body armour": "Body Armour",
  "body armours": "Body Armour",
  boot: "Boots",
  boots: "Boots",
  bow: "Bow",
  bows: "Bow",
  claw: "Claw",
  claws: "Claw",
  contract: "Contract",
  contracts: "Contract",
  dagger: "Dagger",
  daggers: "Dagger",
  flask: "Flask",
  flasks: "Flask",
  glove: "Gloves",
  gloves: "Gloves",
  helmet: "Helmet",
  helmets: "Helmet",
  jewel: "Jewel",
  jewels: "Jewel",
  mace: "Mace",
  maces: "Mace",
  map: "Map",
  maps: "Map",
  quiver: "Quiver",
  quivers: "Quiver",
  ring: "Ring",
  rings: "Ring",
  shield: "Shield",
  shields: "Shield",
  staff: "Staff",
  staves: "Staff",
  sword: "Sword",
  swords: "Sword",
  tincture: "Tincture",
  tinctures: "Tincture",
  wand: "Wand",
  wands: "Wand",
};

const IMPORT_NAME_ALIASES = new Map([
  ["advanced fortress", "advancing fortress"],
]);

const KNOWN_IMPORT_STATUSES = new Set([
  "owned",
  "wearing",
  "owned foil",
  "owned rainbow",
  "missing",
  "unreviewed",
]);

function normalizeImportItemType(sheetName: string) {
  return (
    IMPORT_SHEET_TYPES[
      sheetName.trim().toLowerCase()
    ] ?? null
  );
}

function normalizeImportName(value: string) {
  return value
    .normalize("NFKC")
    .replace(/[’‘]/g, "'")
    .replace(/[^a-zA-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function importNameWithAlias(value: string) {
  const normalized = normalizeImportName(value);
  return IMPORT_NAME_ALIASES.get(normalized) ?? normalized;
}

function importNameDistance(leftValue: string, rightValue: string) {
  const left = importNameWithAlias(leftValue);
  const right = importNameWithAlias(rightValue);

  if (left === right) {
    return 0;
  }

  if (!left) {
    return right.length;
  }

  if (!right) {
    return left.length;
  }

  let previous = Array.from(
    { length: right.length + 1 },
    (_, index) => index,
  );

  for (let leftIndex = 1; leftIndex <= left.length; leftIndex += 1) {
    const current = [leftIndex];

    for (let rightIndex = 1; rightIndex <= right.length; rightIndex += 1) {
      current[rightIndex] = Math.min(
        current[rightIndex - 1] + 1,
        previous[rightIndex] + 1,
        previous[rightIndex - 1] +
          (left[leftIndex - 1] === right[rightIndex - 1] ? 0 : 1),
      );
    }

    previous = current;
  }

  return previous[right.length];
}

function isLikelySameImportName(left: string, right: string) {
  const normalizedLeft = importNameWithAlias(left);
  const normalizedRight = importNameWithAlias(right);

  if (!normalizedLeft || !normalizedRight) {
    return false;
  }

  if (normalizedLeft === normalizedRight) {
    return true;
  }

  const longest = Math.max(
    normalizedLeft.length,
    normalizedRight.length,
  );
  const distance = importNameDistance(
    normalizedLeft,
    normalizedRight,
  );

  if (longest < 8) {
    return distance <= 1;
  }

  if (longest < 20) {
    return distance <= 2;
  }

  return distance <= 3 && distance / longest <= 0.12;
}

function releaseLine(value: string | null | undefined) {
  const match = /^(\d+)\.(\d+)/.exec(value?.trim() ?? "");

  if (!match) {
    return null;
  }

  return `${Number(match[1])}.${Number(match[2])}`;
}

function latestReleaseLine(values: Array<string | null>) {
  let best: { major: number; minor: number; label: string } | null = null;

  for (const value of values) {
    const match = /^(\d+)\.(\d+)/.exec(value?.trim() ?? "");

    if (!match) {
      continue;
    }

    const major = Number(match[1]);
    const minor = Number(match[2]);

    if (
      !best ||
      major > best.major ||
      (major === best.major && minor > best.minor)
    ) {
      best = {
        major,
        minor,
        label: `${major}.${minor}`,
      };
    }
  }

  return best?.label ?? null;
}


const ITEM_TYPE_ORDER = [
  "Flask",
  "Amulet",
  "Ring",
  "Claw",
  "Dagger",
  "Wand",
  "Sword",
  "Axe",
  "Mace",
  "Bow",
  "Staff",
  "Quiver",
  "Belt",
  "Gloves",
  "Boots",
  "Body Armour",
  "Helmet",
  "Shield",
  "Map",
  "Jewel",
  "Contract",
  "Tincture",
];

function compareItemTypes(left: string, right: string) {
  const leftIndex = ITEM_TYPE_ORDER.indexOf(left);
  const rightIndex = ITEM_TYPE_ORDER.indexOf(right);

  const leftOrder =
    leftIndex === -1 ? ITEM_TYPE_ORDER.length : leftIndex;

  const rightOrder =
    rightIndex === -1 ? ITEM_TYPE_ORDER.length : rightIndex;

  return (
    leftOrder - rightOrder ||
    left.localeCompare(right)
  );
}



const DEFAULT_STATUS_COLORS: StatusColors = {
  missing: "#d76868",
  owned: "#79ba68",
  wearing: "#71a5dc",
  foil: "rainbow",
  foulborn: "#c99c72",
  vestigial: "#b79bd7",
};

const STATUS_LABELS: Record<StatusKey, string> = {
  missing: "Missing",
  owned: "Owned",
  wearing: "Wearing",
  foil: "Foil",
  foulborn: "Foulborn",
  vestigial: "Vestigial",
};

const STATUS_KEYS: StatusKey[] = [
  "missing",
  "owned",
  "wearing",
  "foil",
  "foulborn",
  "vestigial",
];

const DEFAULT_COLLECTION_RULES: CollectionRules = {
  owned: true,
  wearing: true,
  foil: true,
  foulborn: true,
  vestigial: true,
};

const DEFAULT_EXTRA_TRACKING: ExtraTracking = {
  foil: true,
  foulborn: true,
  vestigial: true,
};

const EXTRA_VARIANTS: ExtraVariant[] = [
  "foil",
  "foulborn",
  "vestigial",
];

const DEFAULT_EDITION_AVAILABILITY: EditionAvailabilityMap = {
  foil: "unknown",
  foulborn: "unknown",
  vestigial: "unknown",
};

const DEFAULT_EDITION_SOURCES: EditionSourceMap = {
  foil: "default",
  foulborn: "default",
  vestigial: "default",
};

const CONFIDENT_EDITION_SOURCES = new Set([
  "manual",
  "poewiki-foulborn-category",
  "poewiki-foil-confirmed",
  "built-in-vestigial-class-rule",
]);

function isExtraVariant(
  value: string,
): value is ExtraVariant {
  return EXTRA_VARIANTS.includes(
    value as ExtraVariant,
  );
}

function isEditionAvailability(
  value: string,
): value is EditionAvailability {
  return (
    value === "available" ||
    value === "unavailable" ||
    value === "unknown"
  );
}

function shouldShowEdition(
  edition: ExtraVariant,
  availability: EditionAvailabilityMap,
  flags: TrackingFlag[],
  extraTracking: ExtraTracking,
  detectedEdition?:
    | "normal"
    | "foulborn"
    | "vestigial",
) {
  if (!extraTracking[edition]) {
    return false;
  }

  /*
   * Existing player data always wins. Never hide a
   * status somebody already has saved.
   *
   * Likewise, if the item parser literally detects a
   * Foulborn/Vestigial item, that is direct evidence
   * that the edition exists.
   */
  if (
    flags.includes(edition) ||
    detectedEdition === edition
  ) {
    return true;
  }

  /*
   * Unknown stays visible until we have enough evidence
   * to call an edition impossible.
   */
  return availability[edition] !== "unavailable";
}

function isEditionUncertain(
  edition: ExtraVariant,
  availability: EditionAvailabilityMap,
  sources: EditionSourceMap,
  flags: TrackingFlag[],
  detectedEdition?:
    | "normal"
    | "foulborn"
    | "vestigial",
) {
  /*
   * Direct evidence always wins.
   *
   * If the parser literally sees the edition in-game,
   * or the player says they own one, PoE Collector
   * treats that edition as confirmed to exist.
   */
  if (
    detectedEdition === edition ||
    flags.includes(edition)
  ) {
    return false;
  }

  if (
    availability[edition] === "unknown"
  ) {
    return true;
  }

  if (
    availability[edition] === "available"
  ) {
    return !CONFIDENT_EDITION_SOURCES.has(
      sources[edition],
    );
  }

  return false;
}

function isTrackingFlag(value: string): value is TrackingFlag {
  return [
    "owned",
    "wearing",
    "foil",
    "foulborn",
    "vestigial",
  ].includes(value);
}

function statusToFlags(status: string): TrackingFlag[] {
  switch (status) {
    case "Owned":
      return ["owned"];
    case "Wearing":
      return ["wearing"];
    case "Owned Foil":
    case "Owned Rainbow":
      return ["foil"];
    default:
      return [];
  }
}

function getStatusStyle(color: string): CSSProperties | undefined {
  if (color === "rainbow") {
    return undefined;
  }

  return {
    "--status-color": color,
  } as CSSProperties;
}

function makeFamilyId(name: string) {
  const slug = name
    .normalize("NFKD")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");

  return `family:${slug || "unknown"}`;
}

function getPoewikiPageId(
  uniqueId: string,
) {
  const match =
    /^poewiki:(\d+):/.exec(
      uniqueId,
    );

  return match?.[1] ?? null;
}

function extractWikiVariantLabel(
  itemName: string,
  pageTitle: string,
) {
  const escapedName =
    itemName.replace(
      /[.*+?^${}()|[\]\\]/g,
      "\\$&",
    );

  const match =
    pageTitle.match(
      new RegExp(
        `^${escapedName} \\((.+)\\)$`,
        "i",
      ),
    );

  return (
    match?.[1]?.trim() ||
    null
  );
}

const SYNTHESIS_VARIANT_FAMILIES =
  new Set([
    "Circle of Anguish",
    "Circle of Fear",
    "Circle of Guilt",
    "Circle of Nostalgia",
    "Circle of Regret",
    "Garb of the Ephemeral",
    "Mask of the Tribunal",
    "Nebulis",
    "Offering to the Serpent",
    "Perepiteia",
  ]);

function getCuratedVariantLabel(
  itemName: string,
  wikiLabel: string | null,
) {
  const label =
    wikiLabel?.trim() ?? "";

  const normalizedLabel =
    label.toLowerCase();

  /*
   * Old Synthesis-era uniques have a current
   * Synthesised form and a historical Fractured form.
   */
  if (
    SYNTHESIS_VARIANT_FAMILIES.has(
      itemName,
    )
  ) {
    if (
      normalizedLabel ===
      "fractured"
    ) {
      return "Fractured — Legacy";
    }

    if (!label) {
      return "Synthesised";
    }
  }

  if (
    itemName ===
    "Cane of Kulemak"
  ) {
    if (
      normalizedLabel ===
      "variant 1"
    ) {
      return "3 Veiled — Catarina Prefix + Generic Prefix + Suffix";
    }

    if (
      normalizedLabel ===
      "variant 2"
    ) {
      return "3 Veiled — Catarina Prefix + 2 Suffixes";
    }

    if (
      normalizedLabel ===
      "variant 3"
    ) {
      return "4 Veiled — 2 Catarina Prefixes + 2 Suffixes";
    }
  }

  if (itemName === "Impresence") {
    const match =
      /^(chaos|cold|fire|lightning|physical)(,\s*full power)?$/i.exec(
        label,
      );

    if (match) {
      const damageType =
        match[1]
          .charAt(0)
          .toUpperCase() +
        match[1].slice(1);

      return match[2]
        ? `${damageType} — Uber Uber Elder — 2 Curses`
        : `${damageType} — Elder — 1 Curse`;
    }
  }

  if (
    itemName ===
    "Storm's Gift"
  ) {
    if (!label) {
      return "Current — Non-Synthesised";
    }

    if (
      normalizedLabel ===
      "synthesised"
    ) {
      return "Synthesised — Legacy";
    }

    if (
      normalizedLabel ===
      "fractured"
    ) {
      return "Fractured — Legacy";
    }
  }

  if (
    itemName ===
    "The Iron Fortress"
  ) {
    if (!label) {
      return "Current";
    }

    if (
      normalizedLabel === "fated"
    ) {
      return "Fated — Legacy";
    }
  }

  if (
    itemName ===
    "Thread of Hope"
  ) {
    if (!label) {
      return "Variable Ring — Small / Medium / Large / Very Large";
    }

    if (
      normalizedLabel ===
      "massive ring"
    ) {
      return "Massive Ring — Uber Sirus";
    }
  }

  if (itemName === "Winterweave") {
    if (!label) {
      return "Current";
    }

    if (
      normalizedLabel === "fated"
    ) {
      return "Fated — Legacy";
    }
  }

  return wikiLabel;
}

function isCuratedLegacyVariant(
  itemName: string,
  variantLabel: string | null,
) {
  const label =
    variantLabel
      ?.trim()
      .toLowerCase() ?? "";

  /*
   * This also catches our already-curated
   * "... — Legacy" labels on later offline starts.
   */
  if (label.includes("legacy")) {
    return true;
  }

  if (
    SYNTHESIS_VARIANT_FAMILIES.has(
      itemName,
    ) &&
    label === "fractured"
  ) {
    return true;
  }

  if (
    itemName === "Storm's Gift" &&
    (
      label === "fractured" ||
      label === "synthesised"
    )
  ) {
    return true;
  }

  if (
    (
      itemName ===
        "The Iron Fortress" ||
      itemName === "Winterweave"
    ) &&
    label === "fated"
  ) {
    return true;
  }

  return false;
}

async function fetchWikiVariantLabels(
  rows: Array<{
    id: string;
    name: string;
  }>,
) {
  /*
   * Only ask the Wiki about families that actually have
   * multiple catalogue rows. Single-version uniques
   * don't need variant labels.
   */
  const familyCounts =
    new Map<string, number>();

  for (const row of rows) {
    const familyId =
      makeFamilyId(row.name);

    familyCounts.set(
      familyId,
      (familyCounts.get(
        familyId,
      ) ?? 0) + 1,
    );
  }

  const candidates =
    rows.filter(
      (row) =>
        (
          familyCounts.get(
            makeFamilyId(
              row.name,
            ),
          ) ?? 0
        ) > 1,
    );

  const pageIds =
    Array.from(
      new Set(
        candidates
          .map((row) =>
            getPoewikiPageId(
              row.id,
            ),
          )
          .filter(
            (
              pageId,
            ): pageId is string =>
              pageId !== null,
          ),
      ),
    );

  const titleByPageId =
    new Map<
      string,
      string
    >();

  /*
   * MediaWiki accepts page IDs in batches. Keeping this
   * at 50 avoids relying on elevated API limits.
   */
  const BATCH_SIZE = 50;

  for (
    let index = 0;
    index < pageIds.length;
    index += BATCH_SIZE
  ) {
    const batch =
      pageIds.slice(
        index,
        index + BATCH_SIZE,
      );

    const params =
      new URLSearchParams({
        action: "query",
        format: "json",
        formatversion: "2",
        origin: "*",
        pageids:
          batch.join("|"),
      });

    try {
      const response =
        await fetch(
          `https://www.poewiki.net/w/api.php?${params.toString()}`,
          {
            method: "GET",
            headers: {
              Accept:
                "application/json",
            },
          },
        );

      if (!response.ok) {
        console.warn(
          `Variant title lookup failed with HTTP ${response.status}.`,
        );

        continue;
      }

      const payload =
        (await response.json()) as
          WikiPageTitleResponse;

      for (
        const page of
        payload.query?.pages ??
        []
      ) {
        if (
          page.pageid == null ||
          !page.title
        ) {
          continue;
        }

        titleByPageId.set(
          String(page.pageid),
          page.title,
        );
      }
    } catch (error) {
      /*
       * Variant labels are optional metadata.
       * Never stop the local catalogue from loading
       * just because the Wiki is unavailable.
       */
      console.warn(
        "Could not fetch PoE Wiki variant titles:",
        error,
      );
    }
  }

  const labels =
    new Map<
      string,
      string
    >();

  for (const row of candidates) {
    const pageId =
      getPoewikiPageId(
        row.id,
      );

    if (!pageId) {
      continue;
    }

    const pageTitle =
      titleByPageId.get(
        pageId,
      );

    if (!pageTitle) {
      continue;
    }

    const label =
      extractWikiVariantLabel(
        row.name,
        pageTitle,
      );

    if (label) {
      labels.set(
        row.id,
        label,
      );
    }
  }

  return labels;
}

// A league profile is identified by GGG's league ID, not by its position
// in the trade endpoint response. The order may change between requests.

function isExcludedPoE1League(id: string) {
  const normalized = id
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();

  return (
    normalized === "standard" ||
    normalized === "hardcore" ||
    normalized === "ruthless" ||
    normalized === "hardcore ruthless" ||
    normalized === "solo self found" ||
    /^(hc|hardcore)\b/.test(normalized) ||
    /\b(ssf|ruthless)\b/.test(normalized)
  );
}

async function fetchActivePoE1Leagues(): Promise<PoeLeagueDiscovery> {
  const response = await fetch(
    "https://www.pathofexile.com/api/leagues?type=main&limit=50",
    {
      headers: {
        Accept: "application/json",
        "User-Agent": "PoE-Collector/1.1.1",
      },
    },
  );

  if (!response.ok) {
    throw new Error(
      `League lookup failed with HTTP ${response.status}.`,
    );
  }

  // GGG's main-league endpoint returns an array.
  const data = (await response.json()) as PoeTradeLeagueEntry[];

  // Never modify collections based on an incomplete response.
  if (
    !Array.isArray(data) ||
    !data.some((league) => league?.id === "Standard")
  ) {
    throw new Error(
      "GGG returned an incomplete main league list.",
    );
  }

  const now = Date.now();
  const activeLeagues: PoeTradeLeagueEntry[] = [];
  const endedLeagueKeys = new Set<string>();
  const seen = new Set<string>();

  for (const league of data) {
    if (!league || typeof league.id !== "string") {
      throw new Error("GGG returned an invalid league entry.");
    }

    
    
    // Only recognize the current official PC challenge
    // league without Hardcore, SSF or Ruthless rules.
    //
    // GGG's category ID does not necessarily match the
    // league ID (for example, Coven vs Covens).
    if (
      league.realm !== "pc" ||
      isExcludedPoE1League(league.id) ||
      league.event === true ||
      league.category?.current !== true ||
      !Array.isArray(league.rules) ||
      league.rules.length !== 0
    ) {
      continue;
    }



    // Unreliable dates must not trigger automatic rollovers.
    if (
      typeof league.startAt !== "string" ||
      (league.endAt !== null &&
        typeof league.endAt !== "string")
    ) {
      throw new Error(
        `Missing league dates for ${league.id}.`,
      );
    }

    const startsAt = Date.parse(league.startAt);
    const endsAt =
      league.endAt === null
        ? Infinity
        : Date.parse(league.endAt);

    if (
      Number.isNaN(startsAt) ||
      Number.isNaN(endsAt)
    ) {
      throw new Error(
        `Invalid league dates for ${league.id}.`,
      );
    }

    // GGG explicitly identifies this league as ended.
    if (endsAt <= now) {
      endedLeagueKeys.add(league.id);
      continue;
    }

    // Don't create collections for upcoming leagues.
    if (
      startsAt > now ||
      seen.has(league.id)
    ) {
      continue;
    }

    seen.add(league.id);
    activeLeagues.push(league);
  }

  console.log(
    "Detected active official PoE leagues:",
    activeLeagues.map((league) => league.id),
  );

  return {
    activeLeagues,
    endedLeagueKeys,
  };
}

async function initializeCollectionProfiles(
  db: Database,
): Promise<{
  profiles: CollectionProfile[];
  activeProfileId: string;
}> {
  await db.execute(`
    CREATE TABLE IF NOT EXISTS collection_profiles (
      id TEXT PRIMARY KEY NOT NULL,
      name TEXT NOT NULL,
      kind TEXT NOT NULL,
      league_key TEXT,
      is_archived INTEGER NOT NULL DEFAULT 0,
      sort_order INTEGER NOT NULL DEFAULT 0,
      created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    )
  `);

  await db.execute(`
    CREATE TABLE IF NOT EXISTS profile_unique_tracking (
      profile_id TEXT NOT NULL,
      unique_id TEXT NOT NULL,
      flag TEXT NOT NULL,
      PRIMARY KEY (profile_id, unique_id, flag)
    )
  `);

  await db.execute(`
    CREATE TABLE IF NOT EXISTS profile_collection_review (
      profile_id TEXT NOT NULL,
      unique_id TEXT NOT NULL,
      reviewed INTEGER NOT NULL DEFAULT 0,
      PRIMARY KEY (profile_id, unique_id)
    )
  `);

  await db.execute(`
  CREATE TABLE IF NOT EXISTS league_rollovers (
    id TEXT PRIMARY KEY NOT NULL,
    archived_profile_id TEXT NOT NULL,
    old_league_key TEXT,
    old_league_name TEXT NOT NULL,
    new_league_key TEXT NOT NULL,
    new_league_name TEXT NOT NULL,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
  )
`);

await db.execute(`
  CREATE TABLE IF NOT EXISTS league_rollover_changes (
    rollover_id TEXT NOT NULL,
    unique_id TEXT NOT NULL,
    flag TEXT NOT NULL,
    PRIMARY KEY (
      rollover_id,
      unique_id,
      flag
    )
  )
`);


  // Remember official leagues across app restarts,
  // even after GGG removes them from its API.
  await db.execute(`
    CREATE TABLE IF NOT EXISTS verified_official_leagues (
      league_key TEXT PRIMARY KEY NOT NULL
    )
  `);



  await db.execute(
    `
      INSERT OR IGNORE INTO collection_profiles (
        id,
        name,
        kind,
        league_key,
        is_archived,
        sort_order
      )
      VALUES (?, ?, 'standard', NULL, 0, 0)
    `,
    [STANDARD_PROFILE_ID, "Standard"],
  );

  const migrationState = await db.select<
    { value: string }[]
  >(`
    SELECT value
    FROM app_meta
    WHERE key = 'collection_profiles_migrated_v1'
  `);

  if (migrationState.length === 0) {
    // Everything tracked before profiles existed belongs to Standard.
    // The old tables are intentionally left untouched as a safety backup.
    await db.execute(
      `
        INSERT OR IGNORE INTO profile_unique_tracking (
          profile_id,
          unique_id,
          flag
        )
        SELECT ?, unique_id, flag
        FROM unique_tracking
      `,
      [STANDARD_PROFILE_ID],
    );

    await db.execute(
      `
        INSERT OR IGNORE INTO profile_collection_review (
          profile_id,
          unique_id,
          reviewed
        )
        SELECT ?, unique_id, reviewed
        FROM collection_review
      `,
      [STANDARD_PROFILE_ID],
    );

    await db.execute(
      `
        INSERT OR REPLACE INTO app_meta (key, value)
        VALUES ('collection_profiles_migrated_v1', 'yes')
      `,
    );

    await db.execute(
      `
        INSERT OR IGNORE INTO app_meta (key, value)
        VALUES ('active_collection_profile', ?)
      `,
      [STANDARD_PROFILE_ID],
    );
  }

  const profileRows = await db.select<
    {
      id: string;
      name: string;
      kind: string;
      is_archived: number;
    }[]
  >(`
    SELECT
      id,
      name,
      kind,
      is_archived
    FROM collection_profiles
    WHERE is_archived = 0
    ORDER BY sort_order ASC, name ASC
  `);

  const profiles: CollectionProfile[] =
    profileRows
      .filter(
        (
          row,
        ): row is {
          id: string;
          name: string;
          kind: "standard" | "challenge";
          is_archived: number;
        } =>
          row.kind === "standard" ||
          row.kind === "challenge",
      )
      .map((row) => ({
        id: row.id,
        name: row.name,
        kind: row.kind,
        isArchived: row.is_archived === 1,
      }));

  const activeRows = await db.select<
    { value: string }[]
  >(`
    SELECT value
    FROM app_meta
    WHERE key = 'active_collection_profile'
  `);

  const requestedActive =
    activeRows[0]?.value ?? STANDARD_PROFILE_ID;

  const activeProfileId = profiles.some(
    (profile) => profile.id === requestedActive,
  )
    ? requestedActive
    : STANDARD_PROFILE_ID;

  if (activeProfileId !== requestedActive) {
    await db.execute(
      `
        INSERT OR REPLACE INTO app_meta (key, value)
        VALUES ('active_collection_profile', ?)
      `,
      [activeProfileId],
    );
  }

  return {
    profiles,
    activeProfileId,
  };
}

async function seedBuiltInSpecialVariants(db: Database) {
  await db.execute(`
    CREATE TABLE IF NOT EXISTS unique_variant_rules (
      id TEXT PRIMARY KEY NOT NULL,
      variant_id TEXT NOT NULL,
      match_type TEXT NOT NULL,
      match_text TEXT NOT NULL,
      priority INTEGER NOT NULL DEFAULT 100,
      source TEXT NOT NULL DEFAULT 'built-in'
    )
  `);

  const seedState = await db.select<{ value: string }[]>(`
    SELECT value
    FROM app_meta
    WHERE key = 'special_variant_rules_seeded_v3'
  `);

  if (seedState.length > 0) {
    return 0;
  }

  const familyId = makeFamilyId("Ralakesh's Impatience");

  await db.execute(
    `
      INSERT OR IGNORE INTO unique_families (
        id,
        name,
        item_type,
        stash_slot_key,
        source
      )
      VALUES (?, ?, 'Boots', ?, 'built-in-special')
    `,
    [
      familyId,
      "Ralakesh's Impatience",
      "ralakesh's impatience",
    ],
  );

  // Keep any collection state that came from the old spreadsheet/live
  // catalogue, but make it explicit that we do not know which of the three
  // modern Ralakesh variants that old row represents.
  await db.execute(
    `
      UPDATE unique_variants
      SET variant_label = 'Imported copy — variant unknown'
      WHERE
        family_id = ?
        AND name = ?
        AND variant_label IS NULL
        AND source != 'built-in-special'
    `,
    [familyId, "Ralakesh's Impatience"],
  );

  const variants = [
    {
      id: "special:ralakesh-impatience:legacy-pre-3-19",
      label: "Legacy — Pre-3.19 Stationary Charge Generation",
      matchType: "contains-all",
      matchText: [
        "Count as having maximum number of Endurance Charges",
        "Count as having maximum number of Frenzy Charges",
        "Count as having maximum number of Power Charges",
        "Gain a Frenzy, Endurance, or Power Charge once per second while you are Stationary",
      ].join("|||"),
      priority: 300,
      releaseVersion: "3.1.0",
      dropEnabled: 0,
      isLegacyOnly: 1,
      removalVersion: "3.19.0",
    },
    {
      id: "special:ralakesh-impatience:legacy-all-charges",
      label: "Legacy — All Three Charges (pre-3.26)",
      matchType: "contains-all",
      matchText: [
        "Count as having maximum number of Endurance Charges",
        "Count as having maximum number of Frenzy Charges",
        "Count as having maximum number of Power Charges",
      ].join("|||"),
      priority: 200,
      releaseVersion: "3.23.0",
      dropEnabled: 0,
      isLegacyOnly: 1,
      removalVersion: "3.26.0",
    },
    {
      id: "special:ralakesh-impatience:endurance",
      label: "Endurance Charge variant",
      matchType: "contains",
      matchText:
        "Count as having maximum number of Endurance Charges",
      priority: 100,
      releaseVersion: "3.26.0",
      dropEnabled: 1,
      isLegacyOnly: 0,
      removalVersion: null,
    },
    {
      id: "special:ralakesh-impatience:frenzy",
      label: "Frenzy Charge variant",
      matchType: "contains",
      matchText:
        "Count as having maximum number of Frenzy Charges",
      priority: 100,
      releaseVersion: "3.26.0",
      dropEnabled: 1,
      isLegacyOnly: 0,
      removalVersion: null,
    },
    {
      id: "special:ralakesh-impatience:power",
      label: "Power Charge variant",
      matchType: "contains",
      matchText:
        "Count as having maximum number of Power Charges",
      priority: 100,
      releaseVersion: "3.26.0",
      dropEnabled: 1,
      isLegacyOnly: 0,
      removalVersion: null,
    },
  ] as const;

  let added = 0;

  for (const variant of variants) {
    const existing = await db.select<{ id: string }[]>(
      `
        SELECT id
        FROM unique_variants
        WHERE id = ?
      `,
      [variant.id],
    );

    if (existing.length === 0) {
      added += 1;
    }

    await db.execute(
      `
        INSERT OR IGNORE INTO unique_variants (
          id,
          family_id,
          name,
          base_type,
          item_type,
          variant_label,
          release_version,
          drop_enabled,
          drop_restricted,
          is_replica,
          has_legacy_variants,
          is_legacy_only,
          removal_version,
          source
        )
        VALUES (
          ?,
          ?,
          ?,
          'Riveted Boots',
          'Boots',
          ?,
          ?,
          ?,
          0,
          0,
          1,
          ?,
          ?,
          'built-in-special'
        )
      `,
      [
        variant.id,
        familyId,
        "Ralakesh's Impatience",
        variant.label,
        variant.releaseVersion,
        variant.dropEnabled,
        variant.isLegacyOnly,
        variant.removalVersion,
      ],
    );

    await db.execute(
      `
        INSERT OR REPLACE INTO unique_variant_rules (
          id,
          variant_id,
          match_type,
          match_text,
          priority,
          source
        )
        VALUES (?, ?, ?, ?, ?, 'built-in-special')
      `,
      [
        `rule:${variant.id}`,
        variant.id,
        variant.matchType,
        variant.matchText,
        variant.priority,
      ],
    );
  }

  await db.execute(`
    INSERT OR REPLACE INTO app_meta (key, value)
    VALUES ('special_variant_rules_seeded_v3', 'yes')
  `);

  return added;
}

const DEFAULT_POE_LOOKUP_HOTKEY =
  "CommandOrControl+Shift+C";

function formatHotkeyForDisplay(
  hotkey: string,
) {
  return hotkey
    .replace(
      "CommandOrControl",
      "Ctrl",
    )
    .split("+")
    .join(" + ");
}

function normalizeCapturedHotkeyKey(
  key: string,
) {
  const upper =
    key.toUpperCase();

  if (/^[A-Z0-9]$/.test(upper)) {
    return upper;
  }

  if (
    /^F([1-9]|1[0-2])$/.test(
      upper,
    )
  ) {
    return upper;
  }

  return null;
}

const DEV_LEAGUE_ROLLOVER_HOTKEY =
  "CommandOrControl+Shift+F12";

const OVERLAY_LABEL = "poe-overlay";
const OVERLAY_WIDTH = 430;
const OVERLAY_HEIGHT = 330;

type OverlayProfileState = {
  id: string;
  name: string;
  flags: TrackingFlag[];
  reviewed: boolean;
};

type OverlayPayload = {
  trackingProfileId: string;
  trackingProfileName: string;
  uniqueId: string;
  name: string;
  baseType: string | null;
  itemType: string;
  variantLabel: string | null;
  edition: "normal" | "foulborn" | "vestigial";
  profileStates: OverlayProfileState[];
  trackingFlags: TrackingFlag[];
  trackingReviewed: boolean;
  editionAvailability: EditionAvailabilityMap;
  editionSources: EditionSourceMap;
  isLegacyOnly: boolean;
  statusColors: StatusColors;
  collectionRules: CollectionRules;
  extraTracking: ExtraTracking;
};

type OverlayAction =
  | {
      kind: "toggle";
      profileId: string;
      uniqueId: string;
      flag: TrackingFlag;
      enabled: boolean;
    }
  | {
      kind: "missing";
      profileId: string;
      uniqueId: string;
    }
  | {
      kind: "profile";
      profileId: string;
    };

function getOverlayPayloadFromUrl(): OverlayPayload | null {
  const params = new URLSearchParams(window.location.search);
  const encoded = params.get("data");

  if (!encoded) {
    return null;
  }

  try {
    return JSON.parse(
      decodeURIComponent(encoded),
    ) as OverlayPayload;
  } catch (error) {
    console.error("Could not parse overlay payload:", error);
    return null;
  }
}

function OverlayApp() {
  const initialPayload = useMemo(
    () => getOverlayPayloadFromUrl(),
    [],
  );

  const [payload, setPayload] =
  useState<OverlayPayload | null>(initialPayload);

const [overlayMessage, setOverlayMessage] =
  useState("Checking hovered item...");

  const lastMatchedNameRef =
  useRef<string | null>(
    initialPayload?.name ?? null,
  );

const suppressNextBlurRef =
  useRef(false);

function hideOverlayManually() {
  suppressNextBlurRef.current = true;

  void WebviewWindow.getCurrent().hide();
}

useEffect(() => {
  document.documentElement.classList.add(
    "poe-overlay-document",
  );
  document.body.classList.add(
    "poe-overlay-document",
  );

  const currentWindow =
    WebviewWindow.getCurrent();

  let focusCheckTimer: number | undefined;
  let snippingGraceUntil = 0;
  let focusTransitionUntil = 0;
  let destroyed = false;

  const scheduleFocusCheck = (delay: number) => {
    if (destroyed) return;
    if (focusCheckTimer !== undefined) {
      window.clearTimeout(focusCheckTimer);
    }
    focusCheckTimer = window.setTimeout(() => {
      focusCheckTimer = undefined;
      void resolveFocusChange();
    }, delay);
  };

  const resolveFocusChange = async () => {
    if (destroyed) return;
    try {
      const mainWindow = await WebviewWindow.getByLabel("main");
      if (mainWindow && await mainWindow.isFocused()) {
        if (Date.now() < focusTransitionUntil) {
          scheduleFocusCheck(100);
          return;
        }

        const itemName = lastMatchedNameRef.current;
        if (itemName) {
          await emit<string>("poe-overlay-search-item", itemName);
        }
        await currentWindow.hide();
        return;
      }

      const foregroundProcess = await invoke<string>(
        "get_foreground_process_name",
      );
      const process = foregroundProcess.toLowerCase();
      const isPathOfExile = process.includes("pathofexile");
      const isSnippingTool = process === "snippingtool.exe" ||
        process === "screenclippinghost.exe";
      const isWindowsFocusTransition = process === "" ||
        process === "explorer.exe" ||
        process === "applicationframehost.exe" ||
        process === "shellexperiencehost.exe";

      if (isSnippingTool) {
        snippingGraceUntil = Date.now() + 20000;
        scheduleFocusCheck(250);
        return;
      }

      if (snippingGraceUntil > Date.now() && !isPathOfExile) {
        scheduleFocusCheck(250);
        return;
      }

      if (isWindowsFocusTransition && Date.now() < focusTransitionUntil) {
        scheduleFocusCheck(100);
        return;
      }

      await currentWindow.hide();
    } catch (error) {
      console.error("Could not resolve overlay focus change:", error);
      if (!destroyed) void currentWindow.hide();
    }
  };

  const handleBlur = () => {
    if (suppressNextBlurRef.current) {
      suppressNextBlurRef.current = false;
      return;
    }
    focusTransitionUntil = Date.now() + 900;
    scheduleFocusCheck(150);
  };

  let unlistenLoading:
    | (() => void)
    | undefined;

  let unlistenResult:
    | (() => void)
    | undefined;

  let unlistenMessage:
    | (() => void)
    | undefined;

  window.addEventListener("blur", handleBlur);

  void listen<void>(
    "poe-overlay-loading",
    () => {
      lastMatchedNameRef.current = null;
      setPayload(null);
      setOverlayMessage(
        "Checking hovered item...",
      );
    },
  ).then((unlisten) => {
    unlistenLoading = unlisten;
  });

  void listen<OverlayPayload>(
    "poe-overlay-result",
    (event) => {
      lastMatchedNameRef.current =
        event.payload.name;

      setPayload(event.payload);
    },
  ).then((unlisten) => {
    unlistenResult = unlisten;
  });

  void listen<string>(
    "poe-overlay-message",
    (event) => {
      lastMatchedNameRef.current = null;
      setPayload(null);
      setOverlayMessage(event.payload);
    },
  ).then((unlisten) => {
    unlistenMessage = unlisten;
  });

  return () => {
    destroyed = true;
    window.removeEventListener(
      "blur",
      handleBlur,
    );

    if (focusCheckTimer !== undefined) {
      window.clearTimeout(
        focusCheckTimer,
      );
    }

    unlistenLoading?.();
    unlistenResult?.();
    unlistenMessage?.();

    document.documentElement.classList.remove(
      "poe-overlay-document",
    );
    document.body.classList.remove(
      "poe-overlay-document",
    );
  };
}, []);

if (!payload) {
  return (
    <div className="poe-overlay-card overlay-error-card">
      <div className="poe-overlay-header">
        <div className="poe-overlay-title-wrap">
          <span className="poe-overlay-kicker">
            POE COLLECTOR
          </span>
          <h1>Unique Tracker</h1>
        </div>

        <button
          type="button"
          className="poe-overlay-close"
          aria-label="Close overlay"
          onClick={hideOverlayManually}
        >
          ×
        </button>
      </div>

      <span>{overlayMessage}</span>
    </div>
  );
}

  const trackingIsMissing = payload.trackingReviewed &&
    !payload.trackingFlags.some((flag) => payload.collectionRules[flag]);

  const editionLabel = payload.edition === "foulborn"
    ? "FOULBORN"
    : payload.edition === "vestigial"
      ? "VESTIGIAL"
      : null;

  const wornProfiles = payload.profileStates.filter(
    (profile) => profile.flags.includes("wearing"),
  );
  const ownedProfiles = payload.profileStates.filter(
    (profile) => profile.flags.length > 0 &&
      !profile.flags.includes("wearing"),
  );

  const statusHeading = "COLLECTION STATUS";
  const statusValue = wornProfiles.length === 0 && ownedProfiles.length === 0
    ? "Unowned"
    : [
        wornProfiles.length > 0
          ? `Currently worn in ${wornProfiles.map((p) => p.name).join(", ")}`
          : null,
        ownedProfiles.length > 0
          ? `Owned in ${ownedProfiles.map((p) => p.name).join(", ")}`
          : null,
      ].filter((part): part is string => part !== null).join(" • ");

      const showFoil =
  shouldShowEdition(
    "foil",
    payload.editionAvailability,
    payload.trackingFlags,
    payload.extraTracking,
    payload.edition,
  );

const showFoulborn =
  shouldShowEdition(
    "foulborn",
    payload.editionAvailability,
    payload.trackingFlags,
    payload.extraTracking,
    payload.edition,
  );

const showVestigial =
  shouldShowEdition(
    "vestigial",
    payload.editionAvailability,
    payload.trackingFlags,
    payload.extraTracking,
    payload.edition,
  );

  const foilUncertain =
  showFoil &&
  isEditionUncertain(
    "foil",
    payload.editionAvailability,
    payload.editionSources,
    payload.trackingFlags,
    payload.edition,
  );

const foulbornUncertain =
  showFoulborn &&
  isEditionUncertain(
    "foulborn",
    payload.editionAvailability,
    payload.editionSources,
    payload.trackingFlags,
    payload.edition,
  );

const vestigialUncertain =
  showVestigial &&
  isEditionUncertain(
    "vestigial",
    payload.editionAvailability,
    payload.editionSources,
    payload.trackingFlags,
    payload.edition,
  );

const hasUncertainEdition =
  foilUncertain ||
  foulbornUncertain ||
  vestigialUncertain;

  async function changeTrackingProfile(profileId: string) {
    if (!payload) return;
    const selected = payload.profileStates.find((profile) => profile.id === profileId);
    if (!selected) return;

    setPayload((current) => current ? {
      ...current,
      trackingProfileId: selected.id,
      trackingProfileName: selected.name,
      trackingFlags: [...selected.flags],
      trackingReviewed: selected.reviewed,
    } : current);

    await emit<OverlayAction>("poe-overlay-action", {
      kind: "profile",
      profileId,
    });
  }

  async function toggleFlag(flag: TrackingFlag) {
    if (!payload) return;
    const enabled = !payload.trackingFlags.includes(flag);
    const enabledFlags: TrackingFlag[] =
      enabled && flag !== "owned" ? ["owned", flag] : [flag];

    setPayload((current) => {
      if (!current) return current;
      const nextTrackingFlags = enabled
        ? Array.from(new Set([...current.trackingFlags, ...enabledFlags]))
        : current.trackingFlags.filter((existing) => existing !== flag);

      return {
        ...current,
        trackingFlags: nextTrackingFlags,
        trackingReviewed: true,
        profileStates: current.profileStates.map((profile) =>
          profile.id === current.trackingProfileId
            ? { ...profile, flags: nextTrackingFlags, reviewed: true }
            : profile,
        ),
      };
    });

    await emit<OverlayAction>("poe-overlay-action", {
      kind: "toggle",
      profileId: payload.trackingProfileId,
      uniqueId: payload.uniqueId,
      flag,
      enabled,
    });
  }

  async function markMissing() {
    if (!payload) return;
    setPayload((current) => current ? {
      ...current,
      trackingFlags: [],
      trackingReviewed: true,
      profileStates: current.profileStates.map((profile) =>
        profile.id === current.trackingProfileId
          ? { ...profile, flags: [], reviewed: true }
          : profile,
      ),
    } : current);

    await emit<OverlayAction>("poe-overlay-action", {
      kind: "missing",
      profileId: payload.trackingProfileId,
      uniqueId: payload.uniqueId,
    });
  }

  return (
    <div className="poe-overlay-card">
      <div className="poe-overlay-header">
        <div className="poe-overlay-title-wrap">
          <span className="poe-overlay-kicker">
            POE COLLECTOR
          </span>
          <h1>{payload.name}</h1>

          <div className="poe-overlay-meta">
            <span>
              {payload.baseType ?? payload.itemType}
            </span>

            {payload.variantLabel && (
              <span>{payload.variantLabel}</span>
            )}
          </div>
        </div>

        <button
          type="button"
          className="poe-overlay-close"
          aria-label="Close overlay"
          onClick={hideOverlayManually}
        >
          ×
        </button>
      </div>

      <div className="poe-overlay-tags">
        {editionLabel && (
          <span
            className={`poe-overlay-edition ${payload.edition}`}
          >
            {editionLabel}
          </span>
        )}

        {payload.isLegacyOnly && (
          <span className="poe-overlay-legacy">
            LEGACY
          </span>
        )}

        {!payload.trackingReviewed && (
          <span className="poe-overlay-unreviewed">
            UNREVIEWED
          </span>
        )}
      </div>

      <div className="poe-overlay-status-heading">
        <span>{statusHeading}</span>
        <strong>{statusValue}</strong>
      </div>

      <div className="poe-overlay-tracking-heading">
        <span>Record in</span>
        <select
          className="poe-overlay-profile-select"
          value={payload.trackingProfileId}
          onChange={(event) => void changeTrackingProfile(event.target.value)}
        >
          {payload.profileStates.map((profile) => (
            <option key={profile.id} value={profile.id}>
              {profile.name}
            </option>
          ))}
        </select>
      </div>

      <div className="poe-overlay-actions">
        <button
          type="button"
          className={`tracking-badge missing ${
            trackingIsMissing ? "active" : ""
          }`}
          style={getStatusStyle(
            payload.statusColors.missing,
          )}
          onClick={() => void markMissing()}
        >
          Missing
        </button>

        <button
          type="button"
          className={`tracking-badge owned ${
            payload.trackingFlags.includes("owned")
              ? "active"
              : ""
          }`}
          style={getStatusStyle(
            payload.statusColors.owned,
          )}
          onClick={() => void toggleFlag("owned")}
        >
          Normal
        </button>

        <button
          type="button"
          className={`tracking-badge wearing ${
            payload.trackingFlags.includes("wearing")
              ? "active"
              : ""
          }`}
          style={getStatusStyle(
            payload.statusColors.wearing,
          )}
          onClick={() => void toggleFlag("wearing")}
        >
          Wearing
        </button>

        {showFoil && (
  <button
    type="button"
    className={`tracking-badge foil ${
      payload.trackingFlags.includes("foil")
        ? "active"
        : ""
    } ${
      payload.statusColors.foil === "rainbow"
        ? "rainbow-status"
        : ""
    }`}
    style={getStatusStyle(
      payload.statusColors.foil,
    )}
    onClick={() => void toggleFlag("foil")}
  >
    {foilUncertain ? "Foil*" : "Foil"}
  </button>
)}

        {showFoulborn && (
  <button
    type="button"
    className={`tracking-badge foulborn ${
      payload.trackingFlags.includes("foulborn")
        ? "active"
        : ""
    }`}
    style={getStatusStyle(
      payload.statusColors.foulborn,
    )}
    onClick={() => void toggleFlag("foulborn")}
  >
    {foulbornUncertain
      ? "Foulborn*"
      : "Foulborn"}
  </button>
)}

        {showVestigial && (
  <button
    type="button"
    className={`tracking-badge vestigial ${
      payload.trackingFlags.includes("vestigial")
        ? "active"
        : ""
    }`}
    style={getStatusStyle(
      payload.statusColors.vestigial,
    )}
    onClick={() => void toggleFlag("vestigial")}
  >
    {vestigialUncertain
      ? "Vestigial*"
      : "Vestigial"}
  </button>
)}
      </div>

      <p
  className={`poe-overlay-hint ${
    hasUncertainEdition
      ? "uncertain"
      : ""
  }`}
>
  {hasUncertainEdition
    ? "*Data inconclusive. This edition may or may not exist."
    : "Click a status to update your collection."}
</p>
    </div>
  );
}


function wait(milliseconds: number) {
  return new Promise<void>((resolve) => {
    window.setTimeout(resolve, milliseconds);
  });
}

function TrackingBadges({
  flags,
  colors,
  extraTracking,
  editionAvailability,
  editionSources,
  isCollected,
  isMissing,
  onToggle,
  onMarkMissing,
}: {
  flags: TrackingFlag[];
  colors: StatusColors;
  extraTracking: ExtraTracking;
  editionAvailability: EditionAvailabilityMap;
  editionSources: EditionSourceMap;
  isCollected: boolean;
  isMissing: boolean;
  onToggle: (flag: TrackingFlag) => void;
  onMarkMissing: () => void;
}) {
  const showFoil =
    shouldShowEdition(
      "foil",
      editionAvailability,
      flags,
      extraTracking,
    );

  const showFoulborn =
    shouldShowEdition(
      "foulborn",
      editionAvailability,
      flags,
      extraTracking,
    );

  const showVestigial =
    shouldShowEdition(
      "vestigial",
      editionAvailability,
      flags,
      extraTracking,
    );

    const foilUncertain =
  showFoil &&
  isEditionUncertain(
    "foil",
    editionAvailability,
    editionSources,
    flags,
  );

const foulbornUncertain =
  showFoulborn &&
  isEditionUncertain(
    "foulborn",
    editionAvailability,
    editionSources,
    flags,
  );

const vestigialUncertain =
  showVestigial &&
  isEditionUncertain(
    "vestigial",
    editionAvailability,
    editionSources,
    flags,
  );

  const hasAnyExtraTracking =
    showFoil ||
    showFoulborn ||
    showVestigial;

  return (
    <div className="tracking-badges">
      <button
        type="button"
        className={`tracking-badge missing ${isMissing ? "active" : ""}`}
        style={getStatusStyle(colors.missing)}
        onClick={onMarkMissing}
      >
        Missing
      </button>

      <button
        type="button"
        className={`tracking-badge owned ${isCollected ? "active" : ""}`}
        style={getStatusStyle(colors.owned)}
        onClick={() => onToggle("owned")}
      >
        Owned
      </button>

      <button
        type="button"
        className={`tracking-badge wearing ${flags.includes("wearing") ? "active" : ""}`}
        style={getStatusStyle(colors.wearing)}
        onClick={() => onToggle("wearing")}
      >
        Wearing
      </button>

      {hasAnyExtraTracking && (
        <span className="tracking-divider" />
      )}

      {showFoil && (
  <button
    type="button"
    className={`tracking-badge foil ${
      flags.includes("foil") ? "active" : ""
    } ${
      colors.foil === "rainbow"
        ? "rainbow-status"
        : ""
    }`}
    style={getStatusStyle(colors.foil)}
    onClick={() => onToggle("foil")}
  >
    {foilUncertain ? "Foil*" : "Foil"}
  </button>
)}

      {showFoulborn && (
  <button
    type="button"
    className={`tracking-badge foulborn ${
      flags.includes("foulborn")
        ? "active"
        : ""
    }`}
    style={getStatusStyle(
      colors.foulborn,
    )}
    onClick={() =>
      onToggle("foulborn")
    }
  >
    {foulbornUncertain
      ? "Foulborn*"
      : "Foulborn"}
  </button>
)}

      {showVestigial && (
  <button
    type="button"
    className={`tracking-badge vestigial ${
      flags.includes("vestigial")
        ? "active"
        : ""
    }`}
    style={getStatusStyle(
      colors.vestigial,
    )}
    onClick={() =>
      onToggle("vestigial")
    }
  >
    {vestigialUncertain
      ? "Vestigial*"
      : "Vestigial"}
  </button>
)}
    </div>
  );
}

function MainApp() {
  const [uniques, setUniques] = useState<UniqueEntry[]>([]);
  const [sourceFile, setSourceFile] = useState("");
  const [appError, setAppError] = useState("");
  const [isImporting, setIsImporting] = useState(false);
  const [database, setDatabase] = useState<Database | null>(null);
  const [databaseReady, setDatabaseReady] = useState(false);
  const [collectionProfiles, setCollectionProfiles] = useState<
    CollectionProfile[]
  >([]);
  
  const [activeProfileId, setActiveProfileIdState] = useState(
    STANDARD_PROFILE_ID,
  );

  const activeProfileIdRef = useRef(STANDARD_PROFILE_ID);
  const collectionLoadIdRef = useRef(0);
  const [loadedCollectionProfileId, setLoadedCollectionProfileId] =
  useState<string | null>(null);

  
function setActiveProfileId(profileId: string) {
  if (activeProfileIdRef.current !== profileId) {
    // Clear the previous collection while the new one loads.
    setUniques([]);
    setLoadedCollectionProfileId(null);
  }

  activeProfileIdRef.current = profileId;
  setActiveProfileIdState(profileId);
}


  const hotkeyLookupRef = useRef<
    () => Promise<void>
  >(async () => {});

  const registeredLookupHotkeyRef =
    useRef<string | null>(null);

  const [searchTerm, setSearchTerm] = useState("");
  const [statusFilter, setStatusFilter] = useState<StatusFilter>("all");
  const [typeFilter, setTypeFilter] = useState("All");
  const [sortMode, setSortMode] = useState<SortMode>("alphabetical");

  const [settingsOpen, setSettingsOpen] = useState(false);
  const [closePromptOpen, setClosePromptOpen] = useState(false);

  const [snapshotBusy, setSnapshotBusy] = useState(false);
  const [snapshotMessage, setSnapshotMessage] = useState("");
  const [snapshotRestoreError, setSnapshotRestoreError] = useState("");
  const [pendingSnapshotRestore, setPendingSnapshotRestore] =
    useState<TrackerSnapshotPreview | null>(null);
  const [backupMode, setBackupMode] = useState<BackupMode>("automatic");
  const [backupFolder, setBackupFolder] = useState("");
  const [backupPromptOpen, setBackupPromptOpen] = useState(false);
  const [backupSettingsMessage, setBackupSettingsMessage] = useState("");
  const [appVersion, setAppVersion] = useState("");

  useEffect(() => {
    void getVersion().then(setAppVersion).catch((error) => {
      console.warn("Could not read tracker version:", error);
    });
  }, []);

  useEffect(() => {
    let unlistenClose: (() => void) | undefined;

    void listen<void>("main-close-requested", () => {
      setClosePromptOpen(true);
    })
      .then((unlisten) => {
        unlistenClose = unlisten;
      })
      .catch((error) => {
        console.error("Could not listen for close requests:", error);
      });

    return () => {
      unlistenClose?.();
    };
  }, []);
  const [statusColors, setStatusColors] =
    useState<StatusColors>(DEFAULT_STATUS_COLORS);
  const [collectionRules, setCollectionRules] =
    useState<CollectionRules>(DEFAULT_COLLECTION_RULES);
  const [extraTracking, setExtraTracking] =
    useState<ExtraTracking>(DEFAULT_EXTRA_TRACKING);
  const [showLegacyVariants, setShowLegacyVariants] =
    useState(true);

  const [catalogueUpdate, setCatalogueUpdate] =
    useState<CatalogueUpdateSummary | null>(null);
  const [catalogueUpdateOpen, setCatalogueUpdateOpen] =
    useState(false);
  const [catalogueChecking, setCatalogueChecking] =
    useState(false);
  const [catalogueCheckMessage, setCatalogueCheckMessage] =
    useState("Catalogue updates have not been checked yet.");

  const [importMatchSummary, setImportMatchSummary] =
    useState<ImportReconciliationSummary | null>(null);

  const [pendingImport, setPendingImport] =
    useState<PendingImport | null>(null);
  const [pendingImportMode, setPendingImportMode] =
    useState<ImportMode>("status-list");
  const [pendingImportProfileId, setPendingImportProfileId] =
    useState(STANDARD_PROFILE_ID);
  const [missingOnlySummary, setMissingOnlySummary] =
    useState<MissingOnlyInferenceSummary | null>(null);
  const [importDetailsExpanded, setImportDetailsExpanded] =
    useState(false);

  const [
    batchImportReviewOpen,
    setBatchImportReviewOpen,
  ] = useState(false);

  const [parserTestText, setParserTestText] = useState("");
  const [parserTestResult, setParserTestResult] =
    useState<ItemIdentificationResult | null>(null);
  const [parserTesting, setParserTesting] = useState(false);
  const [
  lookupHotkey,
  setLookupHotkey,
] = useState(
  DEFAULT_POE_LOOKUP_HOTKEY,
);

const [
  hotkeyRecording,
  setHotkeyRecording,
] = useState(false);

const [
  hotkeySettingMessage,
  setHotkeySettingMessage,
] = useState("");
  const [hotkeyReady, setHotkeyReady] = useState(false);
  const [hotkeyBusy, setHotkeyBusy] = useState(false);
  const [hotkeyMessage, setHotkeyMessage] = useState(
    "Global hotkey is starting...",
  );

  const [rolloverPreview, setRolloverPreview] =
  useState<LeagueRolloverPreview | null>(null);

    const completedRolloverQueueRef =
    useRef<LeagueRolloverPreview[]>([]);

const [rolloverPreviewOpen, setRolloverPreviewOpen] =
  useState(false);

  const [rolloverMode, setRolloverMode] =
  useState<LeagueRolloverMode>(
    "dev-pending",
  );

const [rolloverApplying, setRolloverApplying] =
  useState(false);


const [newLeagueNotice, setNewLeagueNotice] =
  useState<string | null>(null);



const [
  rolloverChangesExpanded,
  setRolloverChangesExpanded,
] = useState(false);

    function isUniqueCollected(unique: UniqueEntry) {
    return (
      unique.reviewed &&
      unique.flags.some(
        (flag) => collectionRules[flag]
      )
    );
  }

  function isUniqueMissing(unique: UniqueEntry) {
    return (
      unique.reviewed &&
      !isUniqueCollected(unique)
    );
  }

  const activeProfile = useMemo(
    () =>
      collectionProfiles.find(
        (profile) => profile.id === activeProfileId,
      ) ?? null,
    [collectionProfiles, activeProfileId],
  );

    async function initializeLiveLeagues(
    db: Database,
    profiles: CollectionProfile[],
  ): Promise<CollectionProfile[]> {
    const newLeagueNames: string[] = [];

    try {
      const {
        activeLeagues: detectedLeagues,
        endedLeagueKeys,
      } = await fetchActivePoE1Leagues();

            // No challenge leagues means Standard is the only
      // active league. Continue checking for ended leagues.
      const activeLeagueKeys = new Set(
        detectedLeagues.map((league) => league.id),
      );

      const existing = await db.select<
        Array<{
          id: string;
          name: string;
          league_key: string | null;
          sort_order: number;
        }>
      >(`
        SELECT id, name, league_key, sort_order
        FROM collection_profiles
        WHERE kind = 'challenge' AND is_archived = 0
        ORDER BY sort_order ASC, created_at ASC
      `);


      // These league names came from GGG's official
      // main-league API, not the trade-league list.
      const confirmedOfficialKeys = new Set([
        ...detectedLeagues.map((league) => league.id),
        ...endedLeagueKeys,
      ]);

      // Save verification permanently so a league remains
      // identifiable after disappearing from the API.
      for (const leagueKey of confirmedOfficialKeys) {
        await db.execute(
          `
            INSERT OR IGNORE INTO verified_official_leagues (
              league_key
            ) VALUES (?)
          `,
          [leagueKey],
        );
      }

      const verifiedRows = await db.select<
        { league_key: string }[]
      >(
        `SELECT league_key FROM verified_official_leagues`,
      );

      const verifiedLeagueKeys = new Set(
        verifiedRows.map((row) => row.league_key),
      );



      let nextProfiles = [...profiles];
      const representedLeagueKeys = new Set(
        existing
          .map((profile) => profile.league_key)
          .filter((key): key is string => Boolean(key)),
      );

      // Preserve the old special league profile and ALL of its tracking.
      // Only assign it a GGG league key when that choice is unambiguous.
      const legacy = existing.find(
        (profile) =>
          profile.id === CURRENT_LEAGUE_PROFILE_ID &&
          !profile.league_key,
      );

      if (legacy) {
        
        // Only associate a saved collection with a league
        // when its name matches the GGG league ID exactly.
        // Never guess based on how many leagues are active.
        const adopt = detectedLeagues.find(
          (league) => league.id === legacy.name,
        );


        if (adopt && !representedLeagueKeys.has(adopt.id)) {
          await db.execute(
            `
              UPDATE collection_profiles
              SET name = ?, league_key = ?
              WHERE id = ? AND league_key IS NULL
            `,
            [adopt.id, adopt.id, legacy.id],
          );

          representedLeagueKeys.add(adopt.id);
          nextProfiles = nextProfiles.map((profile) =>
            profile.id === legacy.id
              ? { ...profile, name: adopt.id }
              : profile,
          );
        }
      }

      // Discover new league/event profiles without moving, merging, or
      // deleting a single existing collection entry.
      const knownIds = new Set(existing.map((profile) => profile.id));
      let nextSortOrder = Math.max(
        10,
        ...existing.map((profile) => profile.sort_order + 1),
      );


      

      for (const league of detectedLeagues) {
        if (representedLeagueKeys.has(league.id)) {
          continue;
        }

        const baseId = `league:${encodeURIComponent(league.id)}`;
        let newId = baseId;
        let suffix = 1;

        while (knownIds.has(newId)) {
          newId = `${baseId}:${suffix}`;
          suffix += 1;
        }

        // Also avoid colliding with archived profiles not shown above.
        const idRows = await db.select<{ id: string }[]>(
          `SELECT id FROM collection_profiles WHERE id = ?`,
          [newId],
        );

        if (idRows.length > 0) {
          newId = `${baseId}:${Date.now()}:${nextSortOrder}`;
        }

        await db.execute(
          `
            INSERT INTO collection_profiles (
              id, name, kind, league_key, is_archived, sort_order
            ) VALUES (?, ?, 'challenge', ?, 0, ?)
          `,
          [newId, league.id, league.id, String(nextSortOrder)],
        );

        // Already-known uniques are Missing in a fresh league. New wiki
        // discoveries can still show up as Unreviewed later.
        await db.execute(
          `
            INSERT OR IGNORE INTO profile_collection_review (
              profile_id, unique_id, reviewed
            )
            SELECT ?, id, 1
            FROM unique_variants
            WHERE source IN ('poewiki', 'built-in-special')
          `,
          [newId],
        );

        knownIds.add(newId);
        representedLeagueKeys.add(league.id);
        
        nextProfiles.push({
          id: newId,
          name: league.id,
          kind: "challenge",
          isArchived: false,
        });

        // Remember this league so we can announce it.
        newLeagueNames.push(league.id);

        nextSortOrder += 1;
      }



      
      // Find saved leagues that GGG no longer lists as active.
      
      
      
      
      
      // One-time migration for collections created before
      // the 3.30 official league tracking system.
      const legacyMigrationState = await db.select<
        { value: string }[]
      >(
        `
          SELECT value
          FROM app_meta
          WHERE key = 'legacy_league_migration_330'
        `,
      );

      const legacyMigrationPending =
        legacyMigrationState.length === 0;

      const legacyMigrationProfiles = legacyMigrationPending
        ? existing.filter(
            (profile) =>
              profile.name !== "Covens of Umbra" &&
              profile.league_key !== "Covens of Umbra" &&
              !activeLeagueKeys.has(
                profile.league_key ?? profile.name,
              ),
          )
        : [];

      const legacyMigrationProfileIds = new Set(
        legacyMigrationProfiles.map((profile) => profile.id),
      );

      async function markLegacyLeagueMigrationComplete() {
        if (!legacyMigrationPending) {
          return;
        }

        await db.execute(
          `
            INSERT OR REPLACE INTO app_meta (key, value)
            VALUES ('legacy_league_migration_330', 'yes')
          `,
        );
      }

      // Normal automatic rollover detection for verified
      // official leagues, excluding our one-time migration.
      const potentiallyEndedProfiles = existing.filter(
        (profile) =>
          profile.league_key !== null &&
          !legacyMigrationProfileIds.has(profile.id) &&
          verifiedLeagueKeys.has(profile.league_key) &&
          !activeLeagueKeys.has(profile.league_key),
      );


      // Double-check with GGG's trade API before archiving.
      // An incomplete main-league response must not cause
      // an active collection to be merged into Standard.
      let endedProfiles = potentiallyEndedProfiles;


      if (potentiallyEndedProfiles.length > 0) {
        const tradeResponse = await fetch(
          "https://www.pathofexile.com/api/trade/data/leagues",
          {
            headers: {
              Accept: "application/json",
              "User-Agent": "PoE-Collector/1.1.1",
            },
          },
        );

        if (!tradeResponse.ok) {
          throw new Error(
            `Trade league verification failed: HTTP ${tradeResponse.status}.`,
          );
        }

        const tradeData = (await tradeResponse.json()) as {
          result?: Array<{ id: string }>;
        };

        if (
          !Array.isArray(tradeData.result) ||
          !tradeData.result.some(
            (league) => league?.id === "Standard",
          )
        ) {
          throw new Error(
            "GGG returned an incomplete trade league list. Rollover skipped.",
          );
        }

        const tradeLeagueKeys = new Set(
          tradeData.result.map((league) => league.id),
        );

        // Only archive if neither official API lists
        // the previously verified league as active.
        endedProfiles = potentiallyEndedProfiles.filter(
          (profile) =>
            !tradeLeagueKeys.has(profile.league_key!),
        );
      }


      // Include the one-time migration alongside normal
      // automatic league rollovers.
      endedProfiles = [
        ...legacyMigrationProfiles,
        ...endedProfiles,
      ];

      if (endedProfiles.length > 0) {
        try {
          // Save a backup automatically, without asking
          // the user to choose a location.
          const backupFolder = await appConfigDir();

          const backup = await createTrackerBackup(
            db,
            backupFolder,
          );

          const completedRollovers: LeagueRolloverPreview[] = [];

          // Merge every ended league into Standard.
          for (const profile of endedProfiles) {
            const completed = await executeLeagueRollover(
              db,
              profile.id,
              true,
            );

            completedRollovers.push(completed);
          }

                    // Only mark the migration complete after every
          // required rollover has successfully committed.
          await markLegacyLeagueMigrationComplete();

          // Refresh profiles and collection data now that
          // the old league collections are archived.
          const refreshed = await initializeCollectionProfiles(db);

          setActiveProfileId(refreshed.activeProfileId);

          await loadCollectionData(
            db,
            refreshed.activeProfileId,
          );

          setSnapshotMessage(
            `Automatic rollover backup saved: ${backup.fileName}`,
          );

          
          // Display the first completed rollover and
          // queue the remaining notifications.
          completedRolloverQueueRef.current =
            completedRollovers.slice(1);

          setRolloverPreview(completedRollovers[0]);
          setRolloverMode("complete");
          setRolloverChangesExpanded(false);
          setRolloverPreviewOpen(true);


          return refreshed.profiles;
        } catch (error) {
          console.error(
            "Automatic league rollover failed:",
            error,
          );

          setAppError(
            `Automatic league rollover failed: ${
              error instanceof Error
                ? error.message
                : String(error)
            }`,
          );

          const refreshed = await initializeCollectionProfiles(db);

          setActiveProfileId(refreshed.activeProfileId);

          await loadCollectionData(
            db,
            refreshed.activeProfileId,
          );

          return refreshed.profiles;
        }
      }

            // A fresh installation may have no old collections
      // to migrate. Mark the migration complete anyway.
      await markLegacyLeagueMigrationComplete();

      return nextProfiles.sort((left, right) => {
        if (left.kind === "standard") return -1;
        if (right.kind === "standard") return 1;
        return left.name.localeCompare(right.name);
      });
    
    } catch (error) {
      // Offline startup and transient GGG API errors must never prevent
      // an existing local collection from loading normally.
      console.warn("Could not discover PoE leagues:", error);
      return profiles;
    } finally {
      // Announce new leagues only after league discovery
      // and automatic rollovers have finished.
      if (newLeagueNames.length > 0) {
        setNewLeagueNotice(newLeagueNames.join(", "));
      }
    }
  }


  async function buildLeagueRolloverPreview(
    db: Database,
    profileId: string,
    newLeagueName: string,
  ): Promise<LeagueRolloverPreview> {
    const leagueDetails = await db.select<
      Array<{ name: string; kind: string; is_archived: number }>
    >(
      `SELECT name, kind, is_archived FROM collection_profiles WHERE id = ?`,
      [profileId],
    );
    const league = leagueDetails[0];
    if (!league || league.kind !== "challenge" || league.is_archived === 1) {
      throw new Error("An active league collection could not be found.");
    }

    const leagueRows = await db.select<
      Array<{ unique_id: string; flag: string; name: string | null }>
    >(
      `
        SELECT tracking.unique_id, tracking.flag, variants.name
        FROM profile_unique_tracking tracking
        LEFT JOIN unique_variants variants
          ON variants.id = tracking.unique_id
        WHERE tracking.profile_id = ?
      `,
      [profileId],
    );
    const standardRows = await db.select<
      Array<{ unique_id: string; flag: string }>
    >(
      `SELECT unique_id, flag FROM profile_unique_tracking WHERE profile_id = ?`,
      [STANDARD_PROFILE_ID],
    );
    const standardFlags = new Set(
      standardRows.map((row) => `${row.unique_id}|||${row.flag}`),
    );
    const changesByUnique = new Map<string, LeagueRolloverChange>();
    const flagCounts: Record<TrackingFlag, number> = {
      owned: 0,
      wearing: 0,
      foil: 0,
      foulborn: 0,
      vestigial: 0,
    };

    for (const row of leagueRows) {
      if (!isTrackingFlag(row.flag)) continue;
      if (standardFlags.has(`${row.unique_id}|||${row.flag}`)) continue;
      flagCounts[row.flag] += 1;
      const existing = changesByUnique.get(row.unique_id);
      if (existing) {
        if (!existing.addedFlags.includes(row.flag)) {
          existing.addedFlags.push(row.flag);
        }
      } else {
        changesByUnique.set(row.unique_id, {
          uniqueId: row.unique_id,
          name: row.name ?? row.unique_id,
          addedFlags: [row.flag],
        });
      }
    }
    const changes = [...changesByUnique.values()].sort(
      (left, right) => left.name.localeCompare(right.name),
    );
    return {
      profileId,
      oldLeagueName: league.name,
      newLeagueName,
      changedUniques: changes.length,
      flagCounts,
      changes,
    };
  }

  async function executeLeagueRollover(
    db: Database,
    profileId: string,
    commitChanges: boolean,
  ): Promise<LeagueRolloverPreview> {
    const preview = await buildLeagueRolloverPreview(
      db,
      profileId,
      "Standard",
    );
    const leagueRows = await db.select<
      Array<{ name: string; league_key: string | null; kind: string }>
    >(
      `SELECT name, league_key, kind FROM collection_profiles WHERE id = ?`,
      [profileId],
    );
    const league = leagueRows[0];
    if (!league || league.kind !== "challenge") {
      throw new Error("The league to archive could not be found.");
    }

    const rolloverId = `rollover:${Date.now()}:${Math.random().toString(36).slice(2)}`;
    const statements: SqliteTransactionStatement[] = [
      {
        sql: `
          INSERT OR IGNORE INTO profile_unique_tracking (
            profile_id, unique_id, flag
          )
          SELECT ?, unique_id, flag
          FROM profile_unique_tracking
          WHERE profile_id = ?
        `,
        params: [STANDARD_PROFILE_ID, profileId],
      },
      {
        sql: `
          INSERT OR REPLACE INTO profile_collection_review (
            profile_id, unique_id, reviewed
          )
          SELECT ?, unique_id, 1
          FROM profile_unique_tracking
          WHERE profile_id = ?
          GROUP BY unique_id
        `,
        params: [STANDARD_PROFILE_ID, profileId],
      },
      {
        sql: `
          INSERT INTO league_rollovers (
            id, archived_profile_id, old_league_key,
            old_league_name, new_league_key, new_league_name
          ) VALUES (?, ?, ?, ?, ?, ?)
        `,
        params: [
          rolloverId,
          profileId,
          league.league_key ?? league.name,
          league.name,
          STANDARD_PROFILE_ID,
          "Standard",
        ],
      },
    ];

    for (const change of preview.changes) {
      for (const flag of change.addedFlags) {
        statements.push({
          sql: `
            INSERT INTO league_rollover_changes (
              rollover_id, unique_id, flag
            ) VALUES (?, ?, ?)
          `,
          params: [rolloverId, change.uniqueId, flag],
        });
      }
    }
    // Preserve all records in the original league. Archiving only hides it
    // from the active dropdown; other ongoing leagues are never modified.
    statements.push({
      sql: `
        UPDATE collection_profiles
        SET is_archived = 1, sort_order = 100
        WHERE id = ? AND kind = 'challenge' AND is_archived = 0
      `,
      params: [profileId],
    });
    await invoke("execute_sqlite_transaction", {
      statements,
      commit: commitChanges,
    });
    return preview;
  }

  async function previewLeagueRollover() {
    if (!database) return;
    try {
      setAppError("");
      const profile =
        collectionProfiles.find(
          (item) => item.id === activeProfileId && item.kind === "challenge",
        ) ?? collectionProfiles.find((item) => item.kind === "challenge");
      if (!profile) {
        throw new Error("No active league is available to test rollover.");
      }
      const preview = await buildLeagueRolloverPreview(
        database,
        profile.id,
        "Standard",
      );
      setRolloverPreview(preview);
      setRolloverMode("dev-pending");
      setRolloverChangesExpanded(false);
      setRolloverPreviewOpen(true);
      const mainWindow = await WebviewWindow.getByLabel("main");
      if (mainWindow) {
        await mainWindow.show();
        await mainWindow.setFocus();
      }
    } catch (error) {
      console.error("Could not simulate ended league:", error);
      setAppError(error instanceof Error ? error.message : String(error));
    }
  }

  async function confirmDevLeagueRollover() {
    if (!database || !rolloverPreview || rolloverApplying) return;
    try {
      setRolloverApplying(true);
      setAppError("");
      const completed = await executeLeagueRollover(
        database,
        rolloverPreview.profileId,
        false,
      );
      setRolloverPreview(completed);
      setRolloverMode("dev-complete");
      setRolloverChangesExpanded(false);
    } catch (error) {
      console.error("Could not test league rollover:", error);
      setAppError(error instanceof Error ? error.message : String(error));
    } finally {
      setRolloverApplying(false);
    }
  }

  async function confirmLeagueRollover() {
    if (!database || !rolloverPreview || rolloverApplying || snapshotBusy) {
      return;
    }
    try {
      setRolloverApplying(true);
      setAppError("");
      setSnapshotMessage("");

      // The user must save a recovery copy BEFORE the transaction can commit.
      // Cancelling the save dialog leaves every collection unchanged.
      const backup = await exportTrackerSnapshot(database);
      if (!backup) return;

      const completed = await executeLeagueRollover(
        database,
        rolloverPreview.profileId,
        true,
      );
      const profileState = await initializeCollectionProfiles(database);
      setCollectionProfiles(profileState.profiles);
      setActiveProfileId(profileState.activeProfileId);
      await loadCollectionData(database, profileState.activeProfileId);
      setSnapshotMessage(`Pre-rollover backup saved as ${backup.fileName}.`);
      setRolloverPreview(completed);
      setRolloverMode("complete");
      setRolloverChangesExpanded(false);
    } catch (error) {
      console.error("Could not archive ended league:", error);
      setAppError(error instanceof Error ? error.message : String(error));
    } finally {
      setRolloverApplying(false);
    }
  }

async function loadImportReconciliationSummary(
  db: Database,
): Promise<ImportReconciliationSummary> {
  const importedRows = await db.select<
    { total: number }[]
  >(`
    SELECT COUNT(*) AS total
    FROM imported_collection
  `);

  const importedTotal = Number(
    importedRows[0]?.total ?? 0,
  );

  const tableRows = await db.select<
    { name: string }[]
  >(`
    SELECT name
    FROM sqlite_master
    WHERE
      type = 'table'
      AND name = 'import_reconciliation'
  `);

  if (tableRows.length === 0) {
    return reconcileImportedCollection(db);
  }

  const reconciledRows = await db.select<
    { total: number }[]
  >(`
    SELECT COUNT(*) AS total
    FROM import_reconciliation
  `);

  const reconciledTotal = Number(
    reconciledRows[0]?.total ?? 0,
  );

  if (reconciledTotal !== importedTotal) {
    return reconcileImportedCollection(db);
  }

  const statusRows = await db.select<
    {
      status: string;
      total: number;
    }[]
  >(`
    SELECT
      status,
      COUNT(*) AS total
    FROM import_reconciliation
    GROUP BY status
  `);

  const summary: ImportReconciliationSummary = {
    matched: 0,
    ambiguous: 0,
    unmatched: 0,
    skipped: 0,
    total: importedTotal,
  };

  for (const row of statusRows) {
    if (
      row.status === "matched" ||
      row.status === "ambiguous" ||
      row.status === "unmatched" ||
      row.status === "skipped"
    ) {
      summary[row.status] = Number(row.total);
    }
  }

  return summary;
}

  
  async function loadCollectionData(
    db: Database,
    profileId = activeProfileIdRef.current,
  ) {
    // Ignore requests for a collection that is no
    // longer selected.
    if (profileId !== activeProfileIdRef.current) {
      return;
    }

    const loadId = ++collectionLoadIdRef.current;

    const canonicalReady =
      await isCanonicalCatalogueReady(db);


    const catalogueItems = await db.select<
      {
        id: string;
        family_id: string;
        name: string;
        base_type: string | null;
        item_type: string;
        variant_label: string | null;
        is_legacy_only: number;
        source: string;
      }[]
    >(`
      SELECT
        v.id,
        v.family_id,
        v.name,
        v.base_type,
        v.item_type,
        v.variant_label,
        v.is_legacy_only,
        v.source
      FROM unique_variants v
      WHERE
        ${
          canonicalReady
            ? "v.source IN ('poewiki', 'built-in-special')"
            : "1 = 1"
        }
        AND NOT (
          v.source = 'poewiki'
          AND EXISTS (
            SELECT 1
            FROM unique_variants special
            WHERE
              special.family_id = v.family_id
              AND special.source = 'built-in-special'
          )
        )
    `);

    const importedItems = await db.select<
      {
        id: string;
        status: string;
      }[]
    >(`
      SELECT id, status
      FROM imported_collection
    `);

    const savedTracking = await db.select<
      {
        unique_id: string;
        flag: string;
      }[]
    >(
      `
        SELECT unique_id, flag
        FROM profile_unique_tracking
        WHERE profile_id = ?
      `,
      [profileId],
    );

    const savedReview = await db.select<
      {
        unique_id: string;
        reviewed: number;
      }[]
    >(
      `
        SELECT unique_id, reviewed
        FROM profile_collection_review
        WHERE profile_id = ?
      `,
      [profileId],
    );

    const editionRows =
  await db.select<
    {
      variant_id: string;
      edition: string;
      availability: string;
      source: string;
    }[]
  >(`
    SELECT
      variant_id,
      edition,
      availability,
      source
    FROM unique_variant_editions
  `);

    const importedStatusById = new Map<string, string>(
      importedItems.map((item) => [
        item.id,
        item.status,
      ] as [string, string]),
    );

    const flagsByUnique = new Map<string, TrackingFlag[]>();

    for (const row of savedTracking) {
      if (!isTrackingFlag(row.flag)) {
        continue;
      }

      const existing =
        flagsByUnique.get(row.unique_id) ?? [];

      existing.push(row.flag);
      flagsByUnique.set(
        row.unique_id,
        existing,
      );
    }

    const reviewedByUnique = new Map<string, boolean>(
      savedReview.map((row) => [
        row.unique_id,
        row.reviewed === 1,
      ] as [string, boolean]),
    );

    const editionAvailabilityByUnique =
  new Map<
    string,
    EditionAvailabilityMap
  >();

  const editionSourcesByUnique =
  new Map<
    string,
    EditionSourceMap
  >();

for (const row of editionRows) {
  if (
    !isExtraVariant(row.edition) ||
    !isEditionAvailability(
      row.availability,
    )
  ) {
    continue;
  }

  const existing =
    editionAvailabilityByUnique.get(
      row.variant_id,
    ) ?? {
      ...DEFAULT_EDITION_AVAILABILITY,
    };

  existing[row.edition] =
    row.availability;

  editionAvailabilityByUnique.set(
    row.variant_id,
    existing,
  );

  const existingSources =
  editionSourcesByUnique.get(
    row.variant_id,
  ) ?? {
    ...DEFAULT_EDITION_SOURCES,
  };

existingSources[row.edition] =
  row.source;

editionSourcesByUnique.set(
  row.variant_id,
  existingSources,
);
}


    // The selected collection may have changed while
    // the database queries were running.
    if (
      profileId !== activeProfileIdRef.current ||
      loadId !== collectionLoadIdRef.current
    ) {
      return;
    }

    setLoadedCollectionProfileId(profileId);

    setUniques(
      catalogueItems.map((item) => {
        const importStatus =
          importedStatusById.get(item.id) ??
          "Unreviewed";

        return {
          id: item.id,
          familyId: item.family_id,
          name: item.name,
          baseType: item.base_type,
          itemType: item.item_type,
          variantLabel: item.variant_label,
          importStatus,
                    flags:
            flagsByUnique.get(item.id) ??
            [],
          editionAvailability:
  editionAvailabilityByUnique.get(
    item.id,
  ) ?? {
    ...DEFAULT_EDITION_AVAILABILITY,
  },
editionSources:
  editionSourcesByUnique.get(
    item.id,
  ) ?? {
    ...DEFAULT_EDITION_SOURCES,
  },
reviewed:
            reviewedByUnique.get(item.id) ??
            false,
          isLegacyOnly: item.is_legacy_only === 1,
          catalogueSource: item.source,
        };
      }),
    );
  }

  async function syncCatalogueFamiliesAndVariants(
    db: Database,
  ) {
    const catalogueRows = await db.select<
      {
        id: string;
        name: string;
        base_type: string | null;
        item_type: string;
        release_version: string | null;
        drop_enabled: number;
        drop_restricted: number;
        is_replica: number;
        has_legacy_variants: number;
        is_legacy_only: number;
        removal_version: string | null;
        source: string;
      }[]
    >(`
      SELECT
        id,
        name,
        base_type,
        item_type,
        release_version,
        drop_enabled,
        drop_restricted,
        is_replica,
        has_legacy_variants,
        is_legacy_only,
        removal_version,
        source
      FROM unique_catalogue
      WHERE source = 'poewiki'
    `);

        const existingVariantLabelRows =
      await db.select<
        {
          id: string;
          variant_label:
            string | null;
        }[]
      >(`
        SELECT
          id,
          variant_label
        FROM unique_variants
        WHERE source = 'poewiki'
      `);

    const existingVariantLabels =
      new Map(
        existingVariantLabelRows.map(
          (row) => [
            row.id,
            row.variant_label,
          ] as const,
        ),
      );

    const wikiVariantLabels =
      await fetchWikiVariantLabels(
        catalogueRows,
      );


    // Remember which catalogue entries existed before this sync.
    // This lets us tell new uniques apart from new versions
    // of existing uniques.
    const existingCanonicalRows = await db.select<
      {
        id: string;
        family_id: string;
      }[]
    >(`
      SELECT id, family_id
      FROM unique_variants
      WHERE source IN ('poewiki', 'built-in-special')
    `);

    const existingVariantIds = new Set(
      existingCanonicalRows.map((row) => row.id),
    );

    const existingFamilyIds = new Set(
      existingCanonicalRows.map((row) => row.family_id),
    );

    const newlyDiscoveredFamilies: string[] = [];
    const newlyDiscoveredVariants: string[] = [];

    for (const row of catalogueRows) {
      if (existingVariantIds.has(row.id)) {
        continue;
      }

      const familyId = makeFamilyId(row.name);

      if (existingFamilyIds.has(familyId)) {
        // Same unique name, but a new distinct version.
        newlyDiscoveredVariants.push(row.id);
      } else {
        // Completely new unique family.
        newlyDiscoveredFamilies.push(row.id);
      }
    }


    for (const row of catalogueRows) {
      const familyId =
        makeFamilyId(row.name);

              const rawVariantLabel =
  wikiVariantLabels.get(
    row.id,
  ) ??
  existingVariantLabels.get(
    row.id,
  ) ??
  null;

const variantLabel =
  getCuratedVariantLabel(
    row.name,
    rawVariantLabel,
  );

const isLegacyOnly =
  row.is_legacy_only === 1 ||
  isCuratedLegacyVariant(
    row.name,
    rawVariantLabel,
  );

      await db.execute(
        `
          INSERT OR REPLACE INTO unique_families (
            id,
            name,
            item_type,
            stash_slot_key,
            source
          )
          VALUES (?, ?, ?, ?, 'poewiki')
        `,
        [
          familyId,
          row.name,
          row.item_type,
          row.name.toLowerCase(),
        ],
      );

      await db.execute(
        `
          INSERT OR REPLACE INTO unique_variants (
            id,
            family_id,
            name,
            base_type,
            item_type,
            variant_label,
            release_version,
            drop_enabled,
            drop_restricted,
            is_replica,
            has_legacy_variants,
            is_legacy_only,
            removal_version,
            source
          )
                    VALUES (
            ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'poewiki'
          )
        `,
                [
          row.id,
          familyId,
          row.name,
          row.base_type,
          row.item_type,
          variantLabel,
          row.release_version,
          row.drop_enabled,
          row.drop_restricted,
          row.is_replica,
          row.has_legacy_variants,
          isLegacyOnly ? 1 : 0,
          row.removal_version,
        ],
      );
    }

    await db.execute(`
      DELETE FROM unique_variants
      WHERE
        source = 'poewiki'
        AND id NOT IN (
          SELECT id
          FROM unique_catalogue
          WHERE source = 'poewiki'
        )
    `);

    // Apply starting review statuses only to newly discovered IDs.
    // Never overwrite an existing player's review decision.
    //
    // 1 = reviewed, with no ownership flags = Missing
    // 0 = Unreviewed
    for (const [ids, initialReviewed] of [
      [newlyDiscoveredFamilies, 1],
      [newlyDiscoveredVariants, 0],
    ] as const) {
      const batchSize = 200;

      for (
        let index = 0;
        index < ids.length;
        index += batchSize
      ) {
        const batch = ids.slice(index, index + batchSize);
        const placeholders = batch.map(() => "?").join(", ");

        await db.execute(
          `
            INSERT OR IGNORE INTO profile_collection_review (
              profile_id,
              unique_id,
              reviewed
            )
            SELECT
              profiles.id,
              variants.id,
              CASE
                WHEN EXISTS (
                  SELECT 1
                  FROM profile_unique_tracking tracking
                  WHERE tracking.profile_id = profiles.id
                    AND tracking.unique_id = variants.id
                )
                THEN 1
                ELSE ?
              END
            FROM collection_profiles profiles
            CROSS JOIN unique_variants variants
            WHERE profiles.is_archived = 0
              AND variants.id IN (${placeholders})
          `,
          [initialReviewed, ...batch],
        );
      }
    }

    await db.execute(`
      DELETE FROM unique_families
      WHERE id NOT IN (
        SELECT DISTINCT family_id
        FROM unique_variants
      )
    `);
  }

    async function runCatalogueCheck(
    db: Database,
    force = false,
  ) {
    if (catalogueChecking) {
      return;
    }

    try {
      if (!force) {
        const shouldCheck = await shouldCheckCatalogueNow(db);

        if (!shouldCheck) {
          return;
        }
      }

      setCatalogueChecking(true);
      setAppError("");
      setCatalogueCheckMessage("Checking PoE Wiki for catalogue updates...");

      const result = await syncCatalogueFromPoeWiki(db);

      await syncCatalogueFamiliesAndVariants(db);
      await seedBuiltInSpecialVariants(db);

      const reconciliation =
        await reconcileImportedCollection(db);

      setImportMatchSummary(reconciliation);

      await loadCollectionData(db, activeProfileIdRef.current);

      const checkedTime = new Date(result.checkedAt).toLocaleString();
      setCatalogueCheckMessage(
        `Last checked ${checkedTime} • ${result.remoteItems.toLocaleString()} catalogue entries received.`,
      );

      if (result.hasVisibleChanges) {
        setCatalogueUpdate({
          revision: result.revision,
          label: result.label,
          newFamilies: result.newFamilies,
          newVariants: result.newVariants,
          dropDisabled: result.dropDisabled,
          updatedEntries: result.updatedEntries,
        });
        setCatalogueUpdateOpen(true);
      }
    } catch (error) {
      console.error("Could not update catalogue from PoE Wiki:", error);

      const message =
        error instanceof Error ? error.message : String(error);

      setCatalogueCheckMessage(
        `Catalogue check failed: ${message}`,
      );

    } finally {
      setCatalogueChecking(false);
    }
  }

  async function acknowledgeCatalogueUpdate(reviewNewItems: boolean) {
    if (!database || !catalogueUpdate) {
      setCatalogueUpdateOpen(false);
      return;
    }

    try {
      await database.execute(
        `
          INSERT OR REPLACE INTO app_meta (key, value)
          VALUES ('acknowledged_catalogue_revision', ?)
        `,
        [catalogueUpdate.revision],
      );

      setCatalogueUpdateOpen(false);

      if (reviewNewItems) {
        setSearchTerm("");
        setTypeFilter("All");
        setStatusFilter("unreviewed");
      }
    } catch (error) {
      console.error("Could not acknowledge catalogue update:", error);
    }
  }

  async function changeStatusColor(status: StatusKey, color: string) {
    const newColors = {
      ...statusColors,
      [status]: color,
    };

    setStatusColors(newColors);

    if (!database) {
      return;
    }

    try {
      await database.execute(
        `
          INSERT OR REPLACE INTO app_meta (key, value)
          VALUES ('status_colors', ?)
        `,
        [JSON.stringify(newColors)],
      );
    } catch (error) {
      console.error("Could not save status colors:", error);
    }
  }

  async function resetStatusColors() {
    setStatusColors(DEFAULT_STATUS_COLORS);

    if (!database) {
      return;
    }

    try {
      await database.execute(
        `
          INSERT OR REPLACE INTO app_meta (key, value)
          VALUES ('status_colors', ?)
        `,
        [JSON.stringify(DEFAULT_STATUS_COLORS)],
      );
    } catch (error) {
      console.error("Could not reset status colors:", error);
    }
  }

  async function changeCollectionRule(flag: TrackingFlag, enabled: boolean) {
    const newRules = {
      ...collectionRules,
      [flag]: enabled,
    };

    setCollectionRules(newRules);

    if (!database) {
      return;
    }

    try {
      await database.execute(
        `
          INSERT OR REPLACE INTO app_meta (key, value)
          VALUES ('collection_rules', ?)
        `,
        [JSON.stringify(newRules)],
      );
    } catch (error) {
      console.error("Could not save collection rules:", error);
    }
  }

  async function changeExtraTracking(
    variant: ExtraVariant,
    enabled: boolean,
  ) {
    const newTracking = {
      ...extraTracking,
      [variant]: enabled,
    };

    setExtraTracking(newTracking);

    if (!enabled && statusFilter === variant) {
      setStatusFilter("all");
    }

    if (!database) {
      return;
    }

    try {
      await database.execute(
        `
          INSERT OR REPLACE INTO app_meta (key, value)
          VALUES ('extra_tracking', ?)
        `,
        [JSON.stringify(newTracking)],
      );
    } catch (error) {
      console.error("Could not save extra tracking settings:", error);
    }
  }

  async function changeShowLegacyVariants(
    enabled: boolean,
  ) {
    setShowLegacyVariants(enabled);

    if (!database) {
      return;
    }

    try {
      await database.execute(
        `
          INSERT OR REPLACE INTO app_meta (key, value)
          VALUES ('show_legacy_variants', ?)
        `,
        [enabled ? "true" : "false"],
      );
    } catch (error) {
      console.error(
        "Could not save legacy display preference:",
        error,
      );
    }
  }

  async function activateLookupHotkey(
    newHotkey: string,
  ) {
    const currentHotkey =
      registeredLookupHotkeyRef.current;

    if (currentHotkey === newHotkey) {
      setHotkeyReady(true);
      setHotkeyMessage(
        `${formatHotkeyForDisplay(newHotkey)} is ready.`,
      );
      return;
    }

    try {
      /*
       * Hot reload can leave a stale registration owned by this app.
       * unregister() only affects PoE Collector's own registration.
       */
      try {
        await unregister(newHotkey);
      } catch {
        // Fine if PoE Collector did not have it registered.
      }

      /*
       * Register the replacement BEFORE releasing the currently-working
       * shortcut. If Windows rejects it because another app owns it, the
       * player's existing hotkey remains active.
       */
      await register(
        newHotkey,
        (event) => {
          if (event.state === "Released") {
            void hotkeyLookupRef.current();
          }
        },
      );

      if (
        currentHotkey &&
        currentHotkey !== newHotkey
      ) {
        try {
          await unregister(currentHotkey);
        } catch (error) {
          console.warn(
            "Could not unregister the previous lookup hotkey:",
            error,
          );
        }
      }

      registeredLookupHotkeyRef.current =
        newHotkey;

      setHotkeyReady(true);
      setHotkeyMessage(
        `${formatHotkeyForDisplay(newHotkey)} is ready.`,
      );
    } catch (error) {
      setHotkeyReady(
        registeredLookupHotkeyRef.current !== null,
      );

      const previous =
        registeredLookupHotkeyRef.current;

      setHotkeyMessage(
        previous
          ? `${formatHotkeyForDisplay(previous)} is still active. The new shortcut could not be registered.`
          : error instanceof Error
            ? `Hotkey unavailable: ${error.message}`
            : `Hotkey unavailable: ${String(error)}`,
      );

      throw error;
    }
  }

  async function changeLookupHotkey(
    newHotkey: string,
  ) {
    if (!database) {
      return;
    }

    if (newHotkey === lookupHotkey) {
      setHotkeyRecording(false);
      setHotkeySettingMessage("");
      return;
    }

    if (
      import.meta.env.DEV &&
      newHotkey ===
        DEV_LEAGUE_ROLLOVER_HOTKEY
    ) {
      setHotkeySettingMessage(
        "That shortcut is reserved for the development league-rollover test.",
      );
      return;
    }

    const previousHotkey =
      lookupHotkey;

    try {
      setAppError("");
      setHotkeySettingMessage(
        "Checking shortcut...",
      );

      /*
       * Prove the shortcut actually works before persisting it.
       * This prevents a conflicting shortcut from replacing a known-good
       * saved hotkey.
       */
      await activateLookupHotkey(
        newHotkey,
      );

      try {
        await database.execute(
          `
            INSERT OR REPLACE INTO app_meta (
              key,
              value
            )
            VALUES (
              'poe_lookup_hotkey',
              ?
            )
          `,
          [newHotkey],
        );
      } catch (error) {
        /*
         * The new shortcut was activated but could not be saved.
         * Restore the previous one so runtime state and saved state agree.
         */
        await activateLookupHotkey(
          previousHotkey,
        ).catch(
          (restoreError) => {
            console.error(
              "Could not restore the previous lookup hotkey:",
              restoreError,
            );
          },
        );

        throw error;
      }

      setLookupHotkey(newHotkey);
      setHotkeyRecording(false);
      setHotkeySettingMessage(
        `Hotkey changed to ${formatHotkeyForDisplay(newHotkey)}.`,
      );
    } catch (error) {
      console.error(
        "Could not change lookup hotkey:",
        error,
      );

      setHotkeySettingMessage(
        error instanceof Error
          ? `Could not use that shortcut: ${error.message}`
          : `Could not use that shortcut: ${String(error)}`,
      );
    }
  }

function captureLookupHotkey(
  event:
    ReactKeyboardEvent<HTMLButtonElement>,
) {
  event.preventDefault();
  event.stopPropagation();

  if (event.key === "Escape") {
    setHotkeyRecording(false);
    setHotkeySettingMessage("");
    return;
  }

  if (
    [
      "Control",
      "Shift",
      "Alt",
      "Meta",
    ].includes(event.key)
  ) {
    return;
  }

  const key =
    normalizeCapturedHotkeyKey(
      event.key,
    );

  if (!key) {
    setHotkeySettingMessage(
      "Use a letter, number, or F1–F12 as the main key.",
    );
    return;
  }

  const parts: string[] = [];

  if (event.ctrlKey) {
    parts.push(
      "CommandOrControl",
    );
  }

  if (event.altKey) {
    parts.push("Alt");
  }

  if (event.shiftKey) {
    parts.push("Shift");
  }

  const isFunctionKey =
    /^F([1-9]|1[0-2])$/.test(key);

  if (
    parts.length === 0 &&
    !isFunctionKey
  ) {
    setHotkeySettingMessage(
      "Use Ctrl, Alt, or Shift with letters and numbers.",
    );
    return;
  }

  parts.push(key);

  void changeLookupHotkey(
    parts.join("+"),
  );
}

  useEffect(() => {
  if (!database) {
    return;
  }

  void ensureOverlayWindow().catch(
    (error) => {
      console.error(
        "Could not prepare overlay window:",
        error,
      );
    },
  );
}, [database]);

useEffect(() => {
  if (!database) {
    return;
  }

  void activateLookupHotkey(
    lookupHotkey,
  ).catch((error) => {
    console.error(
      "Could not register global shortcut:",
      error,
    );
  });
}, [database, lookupHotkey]);

useEffect(() => {
  if (
    !database ||
    !import.meta.env.DEV
  ) {
    return;
  }

  let cancelled = false;

  void (async () => {
    try {
      try {
        await unregister(
          DEV_LEAGUE_ROLLOVER_HOTKEY,
        );
      } catch {
        // Fine if it was not registered.
      }

      await register(
        DEV_LEAGUE_ROLLOVER_HOTKEY,
        (event) => {
          if (
            !cancelled &&
            event.state === "Released"
          ) {
            void previewLeagueRollover();
          }
        },
      );
    } catch (error) {
      console.error(
        "Could not register development rollover shortcut:",
        error,
      );
    }
  })();

  return () => {
    cancelled = true;

    void unregister(
      DEV_LEAGUE_ROLLOVER_HOTKEY,
    ).catch(() => {});
  };
}, [database]);

useEffect(() => {
  return () => {
    const currentHotkey =
      registeredLookupHotkeyRef.current;

    if (currentHotkey) {
      void unregister(
        currentHotkey,
      ).catch(() => {});
    }
  };
}, []);

  useEffect(() => {
  let unlistenSearch:
    | (() => void)
    | undefined;

  void listen<string>(
    "poe-overlay-search-item",
    (event) => {
      setSearchTerm(event.payload);
      setStatusFilter("all");
      setTypeFilter("All");
    },
  ).then((unlisten) => {
    unlistenSearch = unlisten;
  });

  return () => {
    unlistenSearch?.();
  };
}, []);

  useEffect(() => {
    if (!database) {
      return;
    }

    let unlistenAction:
      | (() => void)
      | undefined;

    void listen<OverlayAction>(
      "poe-overlay-action",
      (event) => {
        const action = event.payload;

        if (action.kind === "profile") {
          void handleCollectionProfileChange(action.profileId);
          return;
        }

        if (action.kind === "missing") {
          void markUniqueMissing(
            action.uniqueId,
            action.profileId,
          );
          return;
        }

        void applyOverlayFlag(
          action.profileId,
          action.uniqueId,
          action.flag,
          action.enabled,
        );
      },
      {
        target: { kind: "Any" },
      },
    ).then((unlisten) => {
      unlistenAction = unlisten;
    });

    return () => {
      unlistenAction?.();
    };
  }, [database, activeProfileId]);

  useEffect(() => {
    async function initializeDatabase() {
      try {
        
        const db = await Database.load("sqlite:poe-collector.db");

        await ensureEditionAvailabilitySchema(
  db,
);

        await db.execute(`
          CREATE TABLE IF NOT EXISTS imported_collection (
            id TEXT PRIMARY KEY NOT NULL,
            name TEXT NOT NULL,
            item_type TEXT NOT NULL,
            status TEXT NOT NULL
          )
        `);

        await db.execute(`
          CREATE TABLE IF NOT EXISTS app_meta (
            key TEXT PRIMARY KEY NOT NULL,
            value TEXT NOT NULL
          )
        `);

        await db.execute(`
          CREATE TABLE IF NOT EXISTS unique_tracking (
            unique_id TEXT NOT NULL,
            flag TEXT NOT NULL,
            PRIMARY KEY (unique_id, flag)
          )
        `);

        await db.execute(`
          CREATE TABLE IF NOT EXISTS unique_catalogue (
            id TEXT PRIMARY KEY NOT NULL,
            name TEXT NOT NULL,
            base_type TEXT,
            item_type TEXT NOT NULL,
            release_version TEXT,
            drop_enabled INTEGER NOT NULL DEFAULT 1,
            drop_restricted INTEGER NOT NULL DEFAULT 0,
            is_replica INTEGER NOT NULL DEFAULT 0,
            has_legacy_variants INTEGER NOT NULL DEFAULT 0,
            is_legacy_only INTEGER NOT NULL DEFAULT 0,
            removal_version TEXT,
            source TEXT NOT NULL DEFAULT 'imported'
          )
        `);

        await db.execute(`
          CREATE TABLE IF NOT EXISTS unique_families (
            id TEXT PRIMARY KEY NOT NULL,
            name TEXT NOT NULL,
            item_type TEXT NOT NULL,
            stash_slot_key TEXT,
            source TEXT NOT NULL DEFAULT 'imported'
          )
        `);

        await db.execute(`
          CREATE TABLE IF NOT EXISTS unique_variants (
            id TEXT PRIMARY KEY NOT NULL,
            family_id TEXT NOT NULL,
            name TEXT NOT NULL,
            base_type TEXT,
            item_type TEXT NOT NULL,
            variant_label TEXT,
            release_version TEXT,
            drop_enabled INTEGER NOT NULL DEFAULT 1,
            drop_restricted INTEGER NOT NULL DEFAULT 0,
            is_replica INTEGER NOT NULL DEFAULT 0,
            has_legacy_variants INTEGER NOT NULL DEFAULT 0,
            is_legacy_only INTEGER NOT NULL DEFAULT 0,
            removal_version TEXT,
            source TEXT NOT NULL DEFAULT 'imported'
          )
        `);

        await db.execute(`
          CREATE TABLE IF NOT EXISTS collection_review (
            unique_id TEXT PRIMARY KEY NOT NULL,
            reviewed INTEGER NOT NULL DEFAULT 0
          )
        `);

        await db.execute(`
          INSERT OR IGNORE INTO collection_review (
            unique_id,
            reviewed
          )
          SELECT id, 1
          FROM imported_collection
        `);

        const addedSpecialVariants =
  await seedBuiltInSpecialVariants(db);

const canonicalReady =
  await isCanonicalCatalogueReady(db);

if (canonicalReady) {
  const reconciliation =
    await loadImportReconciliationSummary(db);

  setImportMatchSummary(reconciliation);
}

        const migrationState = await db.select<{ value: string }[]>(`
          SELECT value
          FROM app_meta
          WHERE key = 'tracking_flags_migrated_v1'
        `);

        if (migrationState.length === 0) {
          await db.execute(`
            INSERT OR IGNORE INTO unique_tracking (unique_id, flag)
            SELECT id, 'owned'
            FROM imported_collection
            WHERE status = 'Owned'
          `);

          await db.execute(`
            INSERT OR IGNORE INTO unique_tracking (unique_id, flag)
            SELECT id, 'wearing'
            FROM imported_collection
            WHERE status = 'Wearing'
          `);

          await db.execute(`
            INSERT OR IGNORE INTO unique_tracking (unique_id, flag)
            SELECT id, 'foil'
            FROM imported_collection
            WHERE status IN ('Owned Foil', 'Owned Rainbow')
          `);

          await db.execute(`
            INSERT OR REPLACE INTO app_meta (key, value)
            VALUES ('tracking_flags_migrated_v1', 'yes')
          `);
        }

        const profileState =
  await initializeCollectionProfiles(db);

setCollectionProfiles(
  profileState.profiles,
);
setActiveProfileId(
  profileState.activeProfileId,
);

        const savedSource = await db.select<{ value: string }[]>(`
          SELECT value
          FROM app_meta
          WHERE key = 'source_file'
        `);

        const savedSortMode = await db.select<{ value: string }[]>(`
          SELECT value
          FROM app_meta
          WHERE key = 'sort_mode'
        `);

        const savedStatusColors = await db.select<{ value: string }[]>(`
          SELECT value
          FROM app_meta
          WHERE key = 'status_colors'
        `);

        const savedCollectionRules = await db.select<{ value: string }[]>(`
          SELECT value
          FROM app_meta
          WHERE key = 'collection_rules'
        `);

        const savedExtraTracking = await db.select<{ value: string }[]>(`
          SELECT value
          FROM app_meta
          WHERE key = 'extra_tracking'
        `);

        const savedShowLegacyVariants = await db.select<{ value: string }[]>(`
          SELECT value
          FROM app_meta
          WHERE key = 'show_legacy_variants'
        `);

        const savedBackupSettings = await db.select<{
          key: string;
          value: string;
        }[]>(`
          SELECT key, value
          FROM app_meta
          WHERE key IN ('restore_backup_mode', 'restore_backup_folder')
        `);

        for (const setting of savedBackupSettings) {
          if (setting.key === 'restore_backup_mode') {
            if (
              setting.value === 'automatic' ||
              setting.value === 'ask' ||
              setting.value === 'off'
            ) {
              setBackupMode(setting.value);
            }
          } else if (setting.key === 'restore_backup_folder') {
            setBackupFolder(setting.value);
          }
        }

        const savedLookupHotkey =
  await db.select<
    { value: string }[]
  >(`
    SELECT value
    FROM app_meta
    WHERE key = 'poe_lookup_hotkey'
  `);

        if (savedStatusColors.length > 0) {
          try {
            const parsedColors = JSON.parse(savedStatusColors[0].value);

            setStatusColors({
              ...DEFAULT_STATUS_COLORS,
              ...parsedColors,
            });
          } catch (error) {
            console.error("Could not load saved status colors:", error);
          }
        }

        if (savedCollectionRules.length > 0) {
          try {
            const parsedRules = JSON.parse(savedCollectionRules[0].value);

            setCollectionRules({
              ...DEFAULT_COLLECTION_RULES,
              ...parsedRules,
            });
          } catch (error) {
            console.error("Could not load collection rules:", error);
          }
        }

        if (savedExtraTracking.length > 0) {
          try {
            const parsedTracking = JSON.parse(savedExtraTracking[0].value);

            setExtraTracking({
              ...DEFAULT_EXTRA_TRACKING,
              ...parsedTracking,
            });
          } catch (error) {
            console.error("Could not load extra tracking settings:", error);
          }
        }

        if (savedShowLegacyVariants.length > 0) {
          setShowLegacyVariants(
            savedShowLegacyVariants[0].value !== "false",
          );
        }

        if (
  savedLookupHotkey.length > 0 &&
  savedLookupHotkey[0].value.trim()
) {
  setLookupHotkey(
    savedLookupHotkey[0].value,
  );
}

        await loadCollectionData(
          db,
          profileState.activeProfileId,
        );

        if (savedSource.length > 0) {
          setSourceFile(savedSource[0].value);
        }

        if (
          savedSortMode.length > 0 &&
          (savedSortMode[0].value === "alphabetical" ||
            savedSortMode[0].value === "type")
        ) {
          setSortMode(savedSortMode[0].value as SortMode);
        }

        

        
        // Finish league detection and any automatic rollover
        // before continuing with catalogue updates.
        await initializeLiveLeagues(
          db,
          profileState.profiles,
        );

        // Reload the collection profiles because an old
        // league may have been archived during rollover.
        const refreshedProfileState =
          await initializeCollectionProfiles(db);

        setCollectionProfiles(
          refreshedProfileState.profiles,
        );

        setActiveProfileId(
          refreshedProfileState.activeProfileId,
        );

        await loadCollectionData(
          db,
          refreshedProfileState.activeProfileId,
        );

        setDatabase(db);
        setDatabaseReady(true);


        if (addedSpecialVariants > 0) {
          setCatalogueUpdate({
            revision: "special-variants-v1",
            label: "Special Variant Catalogue Update",
            newFamilies: 0,
            newVariants: addedSpecialVariants,
            dropDisabled: 0,
            updatedEntries: 0,
          });
          setCatalogueUpdateOpen(true);
        }

        // Open immediately with the local database, then check the live
        // catalogue in the background at most once every 24 hours.
        void runCatalogueCheck(db, false);
      } catch (error) {
        console.error(error);

        setAppError(
          error instanceof Error ? error.message : String(error),
        );
      }
    }

    initializeDatabase();
  }, []);

  const visibleUniques = useMemo(
    () =>
      showLegacyVariants
        ? uniques
        : uniques.filter(
            (unique) => !unique.isLegacyOnly,
          ),
    [uniques, showLegacyVariants],
  );

  const trackingCounts = useMemo(() => {
    return [
            {
        label: "Owned",
        count: visibleUniques.filter(
          (unique) => isUniqueCollected(unique)
        ).length,
      },
      {
        label: "Missing",
        count: visibleUniques.filter((unique) =>
          isUniqueMissing(unique),
        ).length,
      },
      {
        label: "Unreviewed",
        count: visibleUniques.filter((unique) => !unique.reviewed).length,
      },
      {
        label: "Wearing",
        count: visibleUniques.filter((unique) => unique.flags.includes("wearing"))
          .length,
      },
      ...(extraTracking.foil
        ? [
            {
              label: "Foil",
              count: visibleUniques.filter((unique) =>
                unique.flags.includes("foil"),
              ).length,
            },
          ]
        : []),
      ...(extraTracking.foulborn
        ? [
            {
              label: "Foulborn",
              count: visibleUniques.filter((unique) =>
                unique.flags.includes("foulborn"),
              ).length,
            },
          ]
        : []),
      ...(extraTracking.vestigial
        ? [
            {
              label: "Vestigial",
              count: visibleUniques.filter((unique) =>
                unique.flags.includes("vestigial"),
              ).length,
            },
          ]
        : []),
    ];
  }, [visibleUniques, collectionRules, extraTracking]);

  
  const availableTypes = useMemo(() => {
    const types = new Set<string>(
      visibleUniques.map((unique) => unique.itemType),
    );

    return Array.from(types).sort(compareItemTypes);
  }, [visibleUniques]);


  const displayedUniques = useMemo(() => {
    const normalizedSearch = searchTerm.trim().toLowerCase();

    const filtered = visibleUniques.filter((unique) => {
      const searchableText = [
        unique.name,
        unique.baseType ?? "",
        unique.variantLabel ?? "",
      ]
        .join(" ")
        .toLowerCase();

      const matchesSearch =
        normalizedSearch === "" ||
        searchableText.includes(normalizedSearch);

      let matchesStatus = true;

      if (statusFilter === "unreviewed") {
        matchesStatus = !unique.reviewed;
      } else if (statusFilter === "missing") {
        matchesStatus = isUniqueMissing(unique);
      } else if (statusFilter === "owned") {
        matchesStatus = isUniqueCollected(unique);
      } else if (statusFilter !== "all") {
        matchesStatus = unique.flags.includes(statusFilter);
      }

      const matchesType =
        typeFilter === "All" || unique.itemType === typeFilter;

      return matchesSearch && matchesStatus && matchesType;
    });

    return [...filtered].sort((a, b) => {
      
      if (sortMode === "type") {
        const typeComparison = compareItemTypes(
          a.itemType,
          b.itemType,
        );

        if (typeComparison !== 0) {
          return typeComparison;
        }
      }


      const nameComparison = a.name.localeCompare(b.name);

      if (nameComparison !== 0) {
        return nameComparison;
      }

      return (a.variantLabel ?? "").localeCompare(
        b.variantLabel ?? "",
      );
    });
  }, [
    visibleUniques,
    searchTerm,
    statusFilter,
    typeFilter,
    sortMode,
    collectionRules,
  ]);

  const hasUncertainEditionData =
  useMemo(
    () =>
      displayedUniques.some(
        (unique) =>
          EXTRA_VARIANTS.some(
            (edition) =>
              shouldShowEdition(
                edition,
                unique.editionAvailability,
                unique.flags,
                extraTracking,
              ) &&
              isEditionUncertain(
                edition,
                unique.editionAvailability,
                unique.editionSources,
                unique.flags,
              ),
          ),
      ),
    [
      displayedUniques,
      extraTracking,
    ],
  );

  const groupedUniques = useMemo(() => {
    const groups = new Map<string, UniqueEntry[]>();

    for (const unique of displayedUniques) {
      const existing = groups.get(unique.itemType) ?? [];
      existing.push(unique);
      groups.set(unique.itemType, existing);
    }

    return Array.from(groups.entries()).map(([itemType, items]) => ({
      itemType,
      items,
    }));
  }, [displayedUniques]);

  async function ensureOverlayWindow() {
  const existing =
    await WebviewWindow.getByLabel(
      OVERLAY_LABEL,
    );

  if (existing) {
    return existing;
  }

  const overlay = new WebviewWindow(
    OVERLAY_LABEL,
    {
      url: "index.html?overlay=1",
      title: "PoE Collector Lookup",
      width: OVERLAY_WIDTH,
      height: OVERLAY_HEIGHT,
      decorations: false,
      resizable: false,
      maximizable: false,
      minimizable: false,
      alwaysOnTop: true,
      skipTaskbar: true,
      focus: false,
      visible: false,
      shadow: true,
    },
  );

  await wait(100);

  return overlay;
}

async function showOverlayLoading(
  capturedCursor: {
    x: number;
    y: number;
  } | null,
) {
  const overlay =
    await ensureOverlayWindow();

  let monitor = null;

  try {
    if (capturedCursor) {
      monitor = await monitorFromPoint(
        capturedCursor.x,
        capturedCursor.y,
      );
    }
  } catch {
    monitor = null;
  }

  if (!monitor) {
    monitor = await primaryMonitor();
  }

  if (monitor) {
    const workPosition =
      monitor.workArea.position.toLogical(
        monitor.scaleFactor,
      );

    const workSize =
      monitor.workArea.size.toLogical(
        monitor.scaleFactor,
      );

    const cursorLogical = capturedCursor
      ? {
          x:
            capturedCursor.x /
            monitor.scaleFactor,
          y:
            capturedCursor.y /
            monitor.scaleFactor,
        }
      : {
          x:
            workPosition.x +
            workSize.width / 2,
          y:
            workPosition.y +
            workSize.height / 2,
        };

    const gap = 18;

    let x =
      cursorLogical.x + gap;

    let y =
      cursorLogical.y + gap;

    const rightEdge =
      workPosition.x +
      workSize.width;

    const bottomEdge =
      workPosition.y +
      workSize.height;

    if (
      x + OVERLAY_WIDTH >
      rightEdge - 8
    ) {
      x =
        cursorLogical.x -
        OVERLAY_WIDTH -
        gap;
    }

    if (
      y + OVERLAY_HEIGHT >
      bottomEdge - 8
    ) {
      y =
        cursorLogical.y -
        OVERLAY_HEIGHT -
        gap;
    }

    x = Math.max(
      workPosition.x + 8,
      Math.min(
        x,
        rightEdge -
          OVERLAY_WIDTH -
          8,
      ),
    );

    y = Math.max(
      workPosition.y + 8,
      Math.min(
        y,
        bottomEdge -
          OVERLAY_HEIGHT -
          8,
      ),
    );

    await overlay.setPosition(
      new LogicalPosition(x, y),
    );
  }

  await emit("poe-overlay-loading");

  // One frame so "Checking..." paints before
  // the hidden window becomes visible.
  await wait(16);

  await overlay.show();
}

  async function applyOverlayFlag(
    profileId: string,
    uniqueId: string,
    flag: TrackingFlag,
    enabled: boolean,
  ) {
    if (!database) return;
    const enabledFlags: TrackingFlag[] = enabled && flag !== "owned"
      ? ["owned", flag]
      : [flag];

    try {
      await database.execute(`
        INSERT OR REPLACE INTO profile_collection_review
          (profile_id, unique_id, reviewed)
        VALUES (?, ?, 1)
      `, [profileId, uniqueId]);

      if (enabled) {
        for (const enabledFlag of enabledFlags) {
          await database.execute(`
            INSERT OR REPLACE INTO profile_unique_tracking
              (profile_id, unique_id, flag)
            VALUES (?, ?, ?)
          `, [profileId, uniqueId, enabledFlag]);
        }
      } else {
        await database.execute(`
          DELETE FROM profile_unique_tracking
          WHERE profile_id = ? AND unique_id = ? AND flag = ?
        `, [profileId, uniqueId, flag]);
      }

      if (profileId === activeProfileId) {
        setUniques((current) => current.map((item) =>
          item.id !== uniqueId ? item : {
            ...item,
            reviewed: true,
            flags: enabled
              ? Array.from(new Set([...item.flags, ...enabledFlags]))
              : item.flags.filter((existingFlag) => existingFlag !== flag),
          },
        ));
      }
    } catch (error) {
      console.error("Could not apply overlay status:", error);
      setAppError(error instanceof Error ? error.message : String(error));
    }
  }

  async function openItemOverlay(
  result: Extract<
    ItemIdentificationResult,
    { status: "matched" }
  >,
) {
    if (!database) {
      return;
    }

    const localEntry = uniques.find(
      (unique) => unique.id === result.uniqueId,
    );

    if (!localEntry) {
      setHotkeyMessage(
        `Matched ${result.name}, but its collection row could not be loaded.`,
      );
      return;
    }

    const profileStateRows = await database.select<{
      id: string;
      name: string;
      flag: string | null;
      reviewed: number;
    }[]>(`
      SELECT
        profiles.id,
        profiles.name,
        tracking.flag,
        COALESCE(review.reviewed, 0) AS reviewed
      FROM collection_profiles profiles
      LEFT JOIN profile_unique_tracking tracking
        ON tracking.profile_id = profiles.id AND tracking.unique_id = ?
      LEFT JOIN profile_collection_review review
        ON review.profile_id = profiles.id AND review.unique_id = ?
      WHERE profiles.is_archived = 0
      ORDER BY profiles.sort_order ASC, profiles.name ASC
    `, [localEntry.id, localEntry.id]);

    const profileStatesById = new Map<string, OverlayProfileState>();
    for (const row of profileStateRows) {
      let profileState = profileStatesById.get(row.id);
      if (!profileState) {
        profileState = {
          id: row.id,
          name: row.name,
          flags: [],
          reviewed: row.reviewed === 1,
        };
        profileStatesById.set(row.id, profileState);
      }
      if (row.flag && isTrackingFlag(row.flag)) {
        profileState.flags.push(row.flag);
      }
    }
    const profileStates = Array.from(profileStatesById.values());

    const payload: OverlayPayload = {
      trackingProfileId: activeProfileId,
      trackingProfileName:
        activeProfile?.name ?? "Standard",
      uniqueId: localEntry.id,
      name: result.name,
      baseType: result.baseType,
      itemType: result.itemType,
      variantLabel: result.variantLabel,
      edition: result.edition,
      profileStates,
      trackingFlags: [...localEntry.flags],
      trackingReviewed: localEntry.reviewed,
      editionAvailability:
  localEntry.editionAvailability,
editionSources:
  localEntry.editionSources,
isLegacyOnly: localEntry.isLegacyOnly,
      statusColors,
      collectionRules,
      extraTracking,
    };

    const preparedOverlay =
  await ensureOverlayWindow();

await emit<OverlayPayload>(
  "poe-overlay-result",
  payload,
);

await preparedOverlay.show();
await preparedOverlay.setFocus();

return;

  }

  async function identifyHoveredItemFromHotkey() {
    if (!database || hotkeyBusy) {
      return;
    }

    setHotkeyBusy(true);
    setHotkeyMessage("Copying hovered Path of Exile item...");

    let previousClipboard = "";
    let capturedCursor: {
      x: number;
      y: number;
    } | null = null;

    try {
      try {
        const cursor = await cursorPosition();

        capturedCursor = {
          x: cursor.x,
          y: cursor.y,
        };
      } catch {
  capturedCursor = null;
}

try {
  previousClipboard = await readText();
      } catch {
        previousClipboard = "";
      }

      const sentinel =
        `__POE_COLLECTOR_WAITING_${Date.now()}__`;

      await writeText(sentinel);

      // The shortcut callback fires on release. Give Windows a brief moment
      // to finish releasing the physical Ctrl/Shift keys before we inject
      // the ordinary Ctrl+C that Path of Exile understands.
      await wait(35);

      await invoke("send_ctrl_c");

      let clipboardText = sentinel;

      for (let attempt = 0; attempt < 30; attempt += 1) {
        await wait(10);
        clipboardText = await readText();

        if (clipboardText !== sentinel) {
          break;
        }
      }

      await showOverlayLoading(
        capturedCursor,
      );

      if (
        !clipboardText.trim() ||
        clipboardText === sentinel
      ) {
        await writeText(previousClipboard);

        setHotkeyMessage(
          "No item text was copied. Make sure the mouse is hovering a Path of Exile item.",
        );

        await emit<string>(
          "poe-overlay-message",
          "No Path of Exile item was copied.",
        );

        return;
      }

      setParserTestText(clipboardText);

      const result = await identifyUniqueFromClipboard(
        database,
        clipboardText,
      );

      setParserTestResult(result);

      if (result.status === "matched") {
        const editionText =
          result.edition === "normal"
            ? ""
            : ` • ${result.edition}`;

        setHotkeyMessage(
          `Matched ${result.name}${editionText}.`,
        );

        await openItemOverlay(
          result,
        );
      } else {
        setHotkeyMessage(
          "The hotkey copied an item, but PoE Collector could not match it cleanly.",
        );

        await emit<string>(
          "poe-overlay-message",
         "Item copied, but no clean catalogue match was found.",
        );
        const messageOverlay =
          await WebviewWindow.getByLabel(
            OVERLAY_LABEL,
          );

        if (messageOverlay) {
          await messageOverlay.setFocus();
        }
      }
    } catch (error) {
      console.error("Global PoE lookup failed:", error);

      if (previousClipboard) {
        try {
          await writeText(previousClipboard);
        } catch {
          // Restoring the old clipboard is best-effort only.
        }
      }

      setHotkeyMessage(
        error instanceof Error
          ? `Hotkey failed: ${error.message}`
          : `Hotkey failed: ${String(error)}`,
      );
    } finally {
      setHotkeyBusy(false);
    }
  }

  hotkeyLookupRef.current =
    identifyHoveredItemFromHotkey;

  async function readClipboardAndIdentify() {
    if (!database) {
      return;
    }

    try {
      setParserTesting(true);
      setParserTestResult(null);

      const clipboardText = await readText();

      if (!clipboardText.trim()) {
        setParserTestText("");
        setParserTestResult({
          status: "invalid",
          parsed: {
            itemClass: null,
            rarity: null,
            name: null,
            baseType: null,
            canonicalName: null,
            canonicalBaseType: null,
            edition: "normal",
            rawText: "",
          },
          message: "The clipboard does not contain any text.",
        });
        return;
      }

      setParserTestText(clipboardText);

      const result = await identifyUniqueFromClipboard(
        database,
        clipboardText,
      );

      setParserTestResult(result);
    } catch (error) {
      console.error("Could not read and identify clipboard:", error);

      setParserTestResult({
        status: "invalid",
        parsed: {
          itemClass: null,
          rarity: null,
          name: null,
          baseType: null,
          canonicalName: null,
          canonicalBaseType: null,
          edition: "normal",
          rawText: "",
        },
        message:
          error instanceof Error ? error.message : String(error),
      });
    } finally {
      setParserTesting(false);
    }
  }

  async function runParserTest() {
    if (!database || !parserTestText.trim()) {
      return;
    }

    try {
      setParserTesting(true);
      setParserTestResult(null);

      const result = await identifyUniqueFromClipboard(
        database,
        parserTestText,
      );

      setParserTestResult(result);
    } catch (error) {
      console.error("Item parser test failed:", error);

      setParserTestResult({
        status: "invalid",
        parsed: {
          itemClass: null,
          rarity: null,
          name: null,
          baseType: null,
          canonicalName: null,
          canonicalBaseType: null,
          edition: "normal",
          rawText: parserTestText,
        },
        message:
          error instanceof Error ? error.message : String(error),
      });
    } finally {
      setParserTesting(false);
    }
  }

  async function handleCollectionProfileChange(
    profileId: string,
  ) {
    if (!database) {
      return;
    }

    const profileExists =
      collectionProfiles.some(
        (profile) => profile.id === profileId,
      );

    if (!profileExists) {
      return;
    }

    try {
      setAppError("");

      await database.execute(
        `
          INSERT OR REPLACE INTO app_meta (key, value)
          VALUES ('active_collection_profile', ?)
        `,
        [profileId],
      );

      setActiveProfileId(profileId);

      await loadCollectionData(
        database,
        profileId,
      );
    } catch (error) {
      console.error(
        "Could not switch collection profile:",
        error,
      );

      setAppError(
        error instanceof Error
          ? error.message
          : String(error),
      );
    }
  }

  async function handleSortModeChange(newSortMode: SortMode) {
    setSortMode(newSortMode);

    if (!database) {
      return;
    }

    try {
      await database.execute(
        `
          INSERT OR REPLACE INTO app_meta (key, value)
          VALUES ('sort_mode', ?)
        `,
        [newSortMode],
      );
    } catch (error) {
      console.error("Could not save sort preference:", error);
    }
  }

  async function markUniqueMissing(
    uniqueId: string,
    profileId = activeProfileId,
  ) {
    if (!database) {
      return;
    }

    try {
      setAppError("");

      await database.execute(
        `
          INSERT OR REPLACE INTO profile_collection_review (
            profile_id,
            unique_id,
            reviewed
          )
          VALUES (?, ?, 1)
        `,
        [profileId, uniqueId],
      );

      await database.execute(
        `
          DELETE FROM profile_unique_tracking
          WHERE profile_id = ? AND unique_id = ?
        `,
        [profileId, uniqueId],
      );

      if (profileId === activeProfileId) {
        setUniques((current) =>
          current.map((item) =>
            item.id === uniqueId
              ? {
                  ...item,
                  flags: [],
                  reviewed: true,
                }
              : item,
          ),
        );
      }
    } catch (error) {
      console.error(error);

      setAppError(
        error instanceof Error ? error.message : String(error),
      );
    }
  }

  async function toggleTrackingFlag(
    uniqueId: string,
    flag: TrackingFlag,
  ) {
    if (!database) return;
    const unique = uniques.find((item) => item.id === uniqueId);
    if (!unique) return;
    const isActive = unique.flags.includes(flag);
    const enabledFlags: TrackingFlag[] = !isActive && flag !== "owned"
      ? ["owned", flag]
      : [flag];

    try {
      setAppError("");
      await database.execute(`
        INSERT OR REPLACE INTO profile_collection_review
          (profile_id, unique_id, reviewed)
        VALUES (?, ?, 1)
      `, [activeProfileId, uniqueId]);

      if (isActive) {
        await database.execute(`
          DELETE FROM profile_unique_tracking
          WHERE profile_id = ? AND unique_id = ? AND flag = ?
        `, [activeProfileId, uniqueId, flag]);
      } else {
        for (const enabledFlag of enabledFlags) {
          await database.execute(`
            INSERT OR REPLACE INTO profile_unique_tracking
              (profile_id, unique_id, flag)
            VALUES (?, ?, ?)
          `, [activeProfileId, uniqueId, enabledFlag]);
        }
      }

      setUniques((current) => current.map((item) =>
        item.id !== uniqueId ? item : {
          ...item,
          reviewed: true,
          flags: isActive
            ? item.flags.filter((existingFlag) => existingFlag !== flag)
            : Array.from(new Set([...item.flags, ...enabledFlags])),
        },
      ));
    } catch (error) {
      console.error(error);
      setAppError(error instanceof Error ? error.message : String(error));
    }
  }

  async function loadTrackableImportCatalogue(
    db: Database,
  ) {
    return db.select<
      {
        id: string;
        family_id: string;
        name: string;
        item_type: string;
        variant_label: string | null;
        release_version: string | null;
        is_legacy_only: number;
      }[]
    >(`
      SELECT
        v.id,
        v.family_id,
        v.name,
        v.item_type,
        v.variant_label,
        v.release_version,
        v.is_legacy_only
      FROM unique_variants v
      WHERE
        v.source IN ('poewiki', 'built-in-special')
        AND NOT (
          v.source = 'poewiki'
          AND EXISTS (
            SELECT 1
            FROM unique_variants special
            WHERE
              special.family_id = v.family_id
              AND special.source = 'built-in-special'
          )
        )
    `);
  }

  async function buildMissingOnlyPreview(
    db: Database,
    rows: ParsedImportRow[],
  ) {
    const catalogue =
      await loadTrackableImportCatalogue(db);

    const currentReleaseLine = latestReleaseLine(
      catalogue.map((item) => item.release_version),
    );

    const protectedLeagueUniques = currentReleaseLine
      ? catalogue
          .filter(
            (item) =>
              item.is_legacy_only !== 1 &&
              releaseLine(item.release_version) === currentReleaseLine &&
              !rows.some(
                (row) =>
                  row.itemType === item.item_type &&
                  isLikelySameImportName(row.name, item.name),
              ),
          )
          .map((item) => ({
            id: item.id,
            name: item.name,
            itemType: item.item_type,
            variantLabel: item.variant_label,
            releaseVersion: item.release_version,
          }))
          .sort((left, right) =>
            left.name.localeCompare(right.name) ||
            (left.variantLabel ?? "").localeCompare(
              right.variantLabel ?? "",
            ),
          )
      : [];

    return {
      latestReleaseLine: currentReleaseLine,
      protectedLeagueUniques,
    };
  }

  function buildImportedEntries(
    rows: ParsedImportRow[],
    mode: ImportMode,
  ): UniqueEntry[] {
    return rows.map((row) => {
      let normalizedStatus = "Missing";

      if (mode === "status-list") {
        const raw = row.rawStatus.trim();
        const normalized = raw.toLowerCase();

        if (normalized === "owned rainbow") {
          normalizedStatus = "Owned Foil";
        } else if (normalized === "owned foil") {
          normalizedStatus = "Owned Foil";
        } else if (normalized === "owned") {
          normalizedStatus = "Owned";
        } else if (normalized === "wearing") {
          normalizedStatus = "Wearing";
        } else if (normalized === "missing") {
          normalizedStatus = "Missing";
        } else {
          normalizedStatus = "Unreviewed";
        }
      }

      return {
        id: row.id,
        familyId: makeFamilyId(row.name),
        name: row.name,
        baseType: null,
        itemType: row.itemType,
        variantLabel: null,
        importStatus: normalizedStatus,
        flags: statusToFlags(normalizedStatus),
        editionAvailability: {
          ...DEFAULT_EDITION_AVAILABILITY,
        },
        editionSources: {
          ...DEFAULT_EDITION_SOURCES,
        },
        reviewed: true,
        isLegacyOnly: false,
        catalogueSource: "imported",
      };
    });
  }

  async function applyImportedMatchesToProfile(
    db: Database,
    destinationProfileId: string,
  ) {
    const matchedRows = await db.select<
      {
        canonical_variant_id: string;
        import_status: string;
      }[]
    >(`
      SELECT
        reconciliation.canonical_variant_id,
        imported.status AS import_status
      FROM import_reconciliation reconciliation
      JOIN imported_collection imported
        ON imported.id = reconciliation.import_id
      WHERE
        reconciliation.status = 'matched'
        AND reconciliation.canonical_variant_id IS NOT NULL
    `);

    const statements: SqliteTransactionStatement[] = [];

    for (const row of matchedRows) {
      const normalizedStatus = row.import_status.trim();
      const flags = statusToFlags(normalizedStatus);
      const isUnreviewed =
        normalizedStatus.toLowerCase() === "unreviewed";

      statements.push({
        sql: `
          DELETE FROM profile_unique_tracking
          WHERE profile_id = ? AND unique_id = ?
        `,
        params: [
          destinationProfileId,
          row.canonical_variant_id,
        ],
      });

      if (isUnreviewed) {
        statements.push({
          sql: `
            DELETE FROM profile_collection_review
            WHERE profile_id = ? AND unique_id = ?
          `,
          params: [
            destinationProfileId,
            row.canonical_variant_id,
          ],
        });
        continue;
      }

      statements.push({
        sql: `
          INSERT OR REPLACE INTO profile_collection_review (
            profile_id,
            unique_id,
            reviewed
          )
          VALUES (?, ?, 1)
        `,
        params: [
          destinationProfileId,
          row.canonical_variant_id,
        ],
      });

      for (const flag of flags) {
        statements.push({
          sql: `
            INSERT OR REPLACE INTO profile_unique_tracking (
              profile_id,
              unique_id,
              flag
            )
            VALUES (?, ?, ?)
          `,
          params: [
            destinationProfileId,
            row.canonical_variant_id,
            flag,
          ],
        });
      }
    }


    if (statements.length > 0) {
      await invoke("execute_sqlite_transaction", {
        statements,
        commit: true,
      });
    }
  }

  async function applyMissingOnlyInferenceToProfile(
    db: Database,
    importedRows: ParsedImportRow[],
    destinationProfileId: string,
  ): Promise<MissingOnlyInferenceSummary> {
    const catalogue =
      await loadTrackableImportCatalogue(db);

    const currentReleaseLine = latestReleaseLine(
      catalogue.map((item) => item.release_version),
    );

    const familyCounts = new Map<string, number>();

    for (const item of catalogue) {
      familyCounts.set(
        item.family_id,
        (familyCounts.get(item.family_id) ?? 0) + 1,
      );
    }

    const reconciliationRows = await db.select<
      {
        import_id: string;
        canonical_variant_id: string | null;
        status: string;
        candidates_json: string;
        name: string;
        item_type: string;
      }[]
    >(`
      SELECT
        reconciliation.import_id,
        reconciliation.canonical_variant_id,
        reconciliation.status,
        reconciliation.candidates_json,
        imported.name,
        imported.item_type
      FROM import_reconciliation reconciliation
      JOIN imported_collection imported
        ON imported.id = reconciliation.import_id
    `);

    const explicitMissingIds = new Set<string>();
    const protectedIds = new Set<string>();

    for (const row of reconciliationRows) {
      if (
        row.status === "matched" &&
        row.canonical_variant_id
      ) {
        explicitMissingIds.add(
          row.canonical_variant_id,
        );
        continue;
      }

      if (row.status === "ambiguous") {
        try {
          const candidates = JSON.parse(
            row.candidates_json,
          ) as Array<{ id?: string }>;

          for (const candidate of candidates) {
            if (candidate.id) {
              protectedIds.add(candidate.id);
            }
          }
        } catch {
          // An unreadable candidate list should never make us infer ownership.
        }
      }

      if (row.status === "unmatched") {
        for (const item of catalogue) {
          if (
            item.item_type === row.item_type &&
            isLikelySameImportName(
              row.name,
              item.name,
            )
          ) {
            protectedIds.add(item.id);
          }
        }
      }
    }

    for (const item of catalogue) {
      if ((familyCounts.get(item.family_id) ?? 0) > 1) {
        protectedIds.add(item.id);
      }

      // If release metadata is missing, absence is not enough evidence to
      // infer ownership safely.
      if (!releaseLine(item.release_version)) {
        protectedIds.add(item.id);
      }
    }

    const protectedNewLeague = currentReleaseLine
      ? catalogue
          .filter(
            (item) =>
              item.is_legacy_only !== 1 &&
              releaseLine(item.release_version) === currentReleaseLine &&
              !explicitMissingIds.has(item.id) &&
              !importedRows.some(
                (row) =>
                  row.itemType === item.item_type &&
                  isLikelySameImportName(
                    row.name,
                    item.name,
                  ),
              ),
          )
          .map((item) => ({
            id: item.id,
            name: item.name,
            itemType: item.item_type,
            variantLabel: item.variant_label,
            releaseVersion: item.release_version,
          }))
          .sort((left, right) =>
            left.name.localeCompare(right.name) ||
            (left.variantLabel ?? "").localeCompare(
              right.variantLabel ?? "",
            ),
          )
      : [];

    const protectedNewIds = new Set(
      protectedNewLeague.map((item) => item.id),
    );

    const inferredOwnedIds = currentReleaseLine
      ? catalogue
          .filter(
            (item) =>
              item.is_legacy_only !== 1 &&
              !explicitMissingIds.has(item.id) &&
              !protectedIds.has(item.id) &&
              !protectedNewIds.has(item.id) &&
              releaseLine(item.release_version) !== null &&
              releaseLine(item.release_version) !== currentReleaseLine,
          )
          .map((item) => item.id)
      : [];

    const statements: SqliteTransactionStatement[] = [];
    const CHUNK_SIZE = 250;

    for (
      let index = 0;
      index < explicitMissingIds.size;
      index += CHUNK_SIZE
    ) {
      const ids = Array.from(explicitMissingIds).slice(
        index,
        index + CHUNK_SIZE,
      );
      const placeholders = ids.map(() => "?").join(", ");

      statements.push(
        {
          sql: `
            DELETE FROM profile_unique_tracking
            WHERE
              profile_id = ?
              AND unique_id IN (${placeholders})
          `,
          params: [destinationProfileId, ...ids],
        },
        {
          sql: `
            INSERT OR REPLACE INTO profile_collection_review (
              profile_id,
              unique_id,
              reviewed
            )
            SELECT ?, id, 1
            FROM unique_variants
            WHERE id IN (${placeholders})
          `,
          params: [destinationProfileId, ...ids],
        },
      );
    }

    for (
      let index = 0;
      index < inferredOwnedIds.length;
      index += CHUNK_SIZE
    ) {
      const ids = inferredOwnedIds.slice(
        index,
        index + CHUNK_SIZE,
      );
      const placeholders = ids.map(() => "?").join(", ");

      statements.push(
        {
          sql: `
            INSERT OR REPLACE INTO profile_collection_review (
              profile_id,
              unique_id,
              reviewed
            )
            SELECT ?, id, 1
            FROM unique_variants
            WHERE id IN (${placeholders})
          `,
          params: [destinationProfileId, ...ids],
        },
        {
          sql: `
            INSERT OR IGNORE INTO profile_unique_tracking (
              profile_id,
              unique_id,
              flag
            )
            SELECT ?, id, 'owned'
            FROM unique_variants
            WHERE id IN (${placeholders})
          `,
          params: [destinationProfileId, ...ids],
        },
      );
    }

    // Uncertain items need manual review rather than
    // being automatically considered Owned or Missing.
    const unreviewedIds = Array.from(protectedIds).filter(
      (id) => !explicitMissingIds.has(id),
    );

    for (
      let index = 0;
      index < unreviewedIds.length;
      index += CHUNK_SIZE
    ) {
      const ids = unreviewedIds.slice(
        index,
        index + CHUNK_SIZE,
      );

      const placeholders = ids.map(() => "?").join(", ");

      statements.push({
        sql: `
          INSERT OR REPLACE INTO profile_collection_review (
            profile_id,
            unique_id,
            reviewed
          )
          SELECT ?, id, 0
          FROM unique_variants
          WHERE id IN (${placeholders})
            AND NOT EXISTS (
              SELECT 1
              FROM profile_unique_tracking tracking
              WHERE tracking.profile_id = ?
                AND tracking.unique_id = unique_variants.id
            )
        `,
        params: [
          destinationProfileId,
          ...ids,
          destinationProfileId,
        ],
      });
    }

    if (statements.length > 0) {
      await invoke("execute_sqlite_transaction", {
        statements,
        commit: true,
      });
    }

    const destinationProfileName =
      collectionProfiles.find(
        (profile) => profile.id === destinationProfileId,
      )?.name ?? "Selected Collection";

    return {
      inferredOwned: inferredOwnedIds.length,
      explicitMissing: explicitMissingIds.size,
      protectedNewLeague,
      protectedAmbiguous: protectedIds.size,
      latestReleaseLine: currentReleaseLine,
      destinationProfileId,
      destinationProfileName,
    };
  }

  function openImportReviewScreen() {
    setSettingsOpen(false);
    setBatchImportReviewOpen(true);
  }

  async function saveCollection(
    imported: UniqueEntry[],
    fileName: string,
    mode: ImportMode,
    parsedRows: ParsedImportRow[],
    destinationProfileId: string,
  ): Promise<MissingOnlyInferenceSummary | null> {
    if (!database) {
      throw new Error(
        "The local database is not ready yet.",
      );
    }

    const previousImports = await database.select<
      { id: string }[]
    >(`
      SELECT id
      FROM imported_collection
    `);

    for (const previous of previousImports) {
      // These IDs belong to raw import rows, not the canonical catalogue.
      // Canonical tracking uses stable poewiki:/special: IDs.
      await database.execute(
        `
          DELETE FROM unique_tracking
          WHERE unique_id = ?
            AND unique_id NOT IN (
              SELECT id
              FROM unique_variants
              WHERE source IN ('poewiki', 'built-in-special')
            )
        `,
        [previous.id],
      );

      await database.execute(
        `
          DELETE FROM collection_review
          WHERE unique_id = ?
            AND unique_id NOT IN (
              SELECT id
              FROM unique_variants
              WHERE source IN ('poewiki', 'built-in-special')
            )
        `,
        [previous.id],
      );
    }

    await database.execute(
      "DELETE FROM imported_collection",
    );

    await database.execute(`
      DELETE FROM import_reconciliation
    `).catch(() => undefined);

    await database.execute(`
      DELETE FROM import_review_progress
    `).catch(() => undefined);

    await database.execute(`
      DELETE FROM import_manual_resolutions
    `).catch(() => undefined);

    for (const unique of imported) {
      await database.execute(
        `
          INSERT INTO imported_collection (
            id,
            name,
            item_type,
            status
          )
          VALUES (?, ?, ?, ?)
        `,
        [
          unique.id,
          unique.name,
          unique.itemType,
          unique.importStatus,
        ],
      );

      // Raw spreadsheet rows are reference data. Reconciliation copies their
      // state onto canonical IDs where a safe match exists.
      await database.execute(
        `
          INSERT OR REPLACE INTO collection_review (
            unique_id,
            reviewed
          )
          VALUES (?, 1)
        `,
        [unique.id],
      );

      for (const flag of unique.flags) {
        await database.execute(
          `
            INSERT OR REPLACE INTO unique_tracking (
              unique_id,
              flag
            )
            VALUES (?, ?)
          `,
          [unique.id, flag],
        );
      }
    }

    await database.execute(
      `
        INSERT OR REPLACE INTO app_meta (
          key,
          value
        )
        VALUES ('source_file', ?)
      `,
      [fileName],
    );

    await database.execute(
      `
        INSERT OR REPLACE INTO app_meta (
          key,
          value
        )
        VALUES ('source_import_mode', ?)
      `,
      [mode],
    );

    const canonicalReady =
      await isCanonicalCatalogueReady(database);

    if (!canonicalReady) {
      if (mode === "missing-only") {
        throw new Error(
          "Missing-only imports need the canonical catalogue before ownership can be inferred.",
        );
      }

      return null;
    }

    const reconciliation =
      await reconcileImportedCollection(
        database,
      );

    setImportMatchSummary(reconciliation);

    await applyImportedMatchesToProfile(
      database,
      destinationProfileId,
    );

    if (mode === "missing-only") {
      return applyMissingOnlyInferenceToProfile(
        database,
        parsedRows,
        destinationProfileId,
      );
    }

    return null;
  }

  async function confirmPendingImport() {
    if (!pendingImport || !database || isImporting) {
      return;
    }

    try {
      setIsImporting(true);
      setAppError("");

      const imported = buildImportedEntries(
        pendingImport.rows,
        pendingImportMode,
      );

      const summary = await saveCollection(
        imported,
        pendingImport.fileName,
        pendingImportMode,
        pendingImport.rows,
        pendingImportProfileId,
      );

      if (pendingImportProfileId === activeProfileId) {
        await loadCollectionData(
          database,
          activeProfileId,
        );
      }

      setSourceFile(pendingImport.fileName);
      setPendingImport(null);
      setImportDetailsExpanded(false);

      if (summary) {
        setMissingOnlySummary(summary);
      }
    } catch (error) {
      console.error(error);
      setAppError(
        error instanceof Error
          ? error.message
          : String(error),
      );
    } finally {
      setIsImporting(false);
    }
  }

  async function handleExportTrackerSnapshot() {
  if (!database || snapshotBusy) {
    return;
  }

  try {
    setSnapshotBusy(true);
    setSnapshotMessage("");
    setAppError("");

    const result = await exportTrackerSnapshot(database);

    if (!result) {
      return;
    }

    setSnapshotMessage(
      `Saved ${result.profileCount} collections to ${result.fileName}.`,
    );
  } catch (error) {
    console.error("Could not export tracker snapshot:", error);

    setAppError(
      error instanceof Error
        ? error.message
        : String(error),
    );
  } finally {
    setSnapshotBusy(false);
  }
}

  async function handleChooseTrackerSnapshot() {
    if (!database || snapshotBusy) return;

    try {
      setSnapshotBusy(true);
      setSnapshotRestoreError("");
      setBackupPromptOpen(false);
      setAppError("");

      const preview = await chooseTrackerSnapshot();
      if (preview) setPendingSnapshotRestore(preview);
    } catch (error) {
      console.error("Could not preview tracker snapshot:", error);
      setAppError(error instanceof Error ? error.message : String(error));
    } finally {
      setSnapshotBusy(false);
    }
  }

  async function saveBackupMode(mode: BackupMode) {
    if (!database) return;
    try {
      await database.execute(
        `INSERT OR REPLACE INTO app_meta (key, value)
         VALUES ('restore_backup_mode', ?)`,
        [mode],
      );
      setBackupMode(mode);
      setBackupSettingsMessage("Backup preference saved.");
      setAppError("");
    } catch (error) {
      setAppError(error instanceof Error ? error.message : String(error));
    }
  }

  async function chooseBackupFolder(): Promise<string | null> {
    if (!database) throw new Error("The local database is not ready yet.");

    const selected = await open({ directory: true, multiple: false });
    if (!selected || Array.isArray(selected)) return null;

    await database.execute(
      `INSERT OR REPLACE INTO app_meta (key, value)
       VALUES ('restore_backup_folder', ?)`,
      [selected],
    );
    setBackupFolder(selected);
    setBackupSettingsMessage("Backup folder saved.");
    return selected;
  }

  async function handleChooseBackupFolder() {
    try {
      setAppError("");
      await chooseBackupFolder();
    } catch (error) {
      setAppError(error instanceof Error ? error.message : String(error));
    }
  }

  // The confirmation always precedes any replacement of collection data.
  // "automatic" writes a backup without prompting if a folder was chosen;
  // "ask" presents a separate decision; "off" is an explicit opt-out.
  async function handleConfirmTrackerSnapshotRestore(
    backupDecision?: boolean,
  ) {
    if (!database || !pendingSnapshotRestore || snapshotBusy) return;

    if (backupMode === "ask" && backupDecision === undefined) {
      setBackupPromptOpen(true);
      return;
    }

    const shouldBackUp = backupDecision ?? backupMode === "automatic";

    try {
      setSnapshotBusy(true);
      setSnapshotRestoreError("");
      setSnapshotMessage("");
      setAppError("");

      let backupPath = "";
      if (shouldBackUp) {
        let folder = backupFolder;
        if (!folder) {
          const selected = await chooseBackupFolder();
          if (!selected) {
            setSnapshotRestoreError(
              "Restore cancelled: a backup folder was not selected.",
            );
            return;
          }
          folder = selected;
        }

        // If creation fails, restore does NOT run.
        const backup = await createTrackerBackup(database, folder);
        backupPath = backup.path;
      }

      const result = await restoreTrackerSnapshot(
        database,
        pendingSnapshotRestore,
      );
      const profileState = await initializeCollectionProfiles(database);

      setCollectionProfiles(profileState.profiles);
      setActiveProfileId(profileState.activeProfileId);
      await loadCollectionData(database, profileState.activeProfileId);
      setSourceFile(result.fileName);
      setPendingSnapshotRestore(null);
      setBackupPromptOpen(false);
      setRolloverPreviewOpen(false);

      setSnapshotMessage(
        `Restored ${result.profileCount} collections from ${result.fileName}.` +
        (backupPath ? ` Previous tracker backed up to ${backupPath}.` : "") +
        (result.skippedUnknownUniques > 0
          ? ` ${result.skippedUnknownUniques} unknown catalogue rows were skipped.`
          : ""),
      );
    } catch (error) {
      console.error("Could not restore tracker snapshot:", error);
      setSnapshotRestoreError(
        error instanceof Error ? error.message : String(error),
      );
    } finally {
      setSnapshotBusy(false);
    }
  }

  async function handleImport() {
    if (!database) {
      setAppError("The local database is not ready yet.");
      return;
    }

    try {
      setAppError("");
      setMissingOnlySummary(null);
      setImportDetailsExpanded(false);

      const selected = await open({
        multiple: false,
        directory: false,
        filters: [
          {
            name: "Excel Workbook",
            extensions: ["xlsx", "xls"],
          },
        ],
      });

      if (!selected || Array.isArray(selected)) {
        return;
      }

      setIsImporting(true);

      const fileBytes = await readFile(selected);

      const workbook = read(fileBytes, {
        type: "array",
        cellFormula: false,
        cellHTML: false,
      });

      const parsedRows: ParsedImportRow[] = [];
      let missingOnlyScore = 0;
      let statusListScore = 0;

      for (const sheetName of workbook.SheetNames) {
        const itemType = normalizeImportItemType(sheetName);

        // Auxiliary sheets such as Ole's ALT sheet are intentionally ignored.
        // Only sheets that map to a real PoE Collector item type are imported.
        if (!itemType) {
          continue;
        }

        const sheet = workbook.Sheets[sheetName];

        if (!sheet) {
          continue;
        }

        const rows = utils.sheet_to_json<
          (string | number | boolean)[]
        >(sheet, {
          header: 1,
          defval: "",
          raw: false,
        });

        rows.forEach((row, rowIndex) => {
          const name = String(row[0] ?? "").trim();
          const secondColumn = String(row[1] ?? "").trim();
          const thirdColumn = String(row[2] ?? "").trim();

          if (rowIndex === 0) {
            const firstHeader = normalizeImportName(name).replace(/:$/, "");
            const secondHeader = normalizeImportName(secondColumn).replace(/:$/, "");
            const thirdHeader = normalizeImportName(thirdColumn).replace(/:$/, "");

            if (
              secondHeader.includes("price") ||
              thirdHeader.includes("acquisition")
            ) {
              missingOnlyScore += 4;
            }

            if (secondHeader.includes("status")) {
              statusListScore += 4;
            }

            if (
              firstHeader === "name" ||
              firstHeader === "item name"
            ) {
              return;
            }
          }

          if (!name) {
            return;
          }

          if (
            KNOWN_IMPORT_STATUSES.has(
              secondColumn.toLowerCase(),
            )
          ) {
            statusListScore += 2;
          }

          parsedRows.push({
            id: `${sheetName}-${rowIndex}-${name}`,
            name,
            itemType,
            rawStatus: secondColumn,
          });
        });
      }

      if (parsedRows.length === 0) {
        throw new Error(
          "No recognizable unique-list rows were found in that workbook.",
        );
      }

      const fileName =
        selected.split(/[\\/]/).pop() ?? selected;

      const preview = await buildMissingOnlyPreview(
        database,
        parsedRows,
      );

      const suggestedMode: ImportMode =
        missingOnlyScore > statusListScore
          ? "missing-only"
          : "status-list";

      setPendingImport({
        fileName,
        rows: parsedRows,
        suggestedMode,
        latestReleaseLine: preview.latestReleaseLine,
        protectedLeagueUniques:
          preview.protectedLeagueUniques,
      });
      setPendingImportMode(suggestedMode);
      setPendingImportProfileId(
        collectionProfiles.some(
          (profile) => profile.id === activeProfileId,
        )
          ? activeProfileId
          : STANDARD_PROFILE_ID,
      );
    } catch (error) {
      console.error(error);

      setAppError(
        error instanceof Error
          ? error.message
          : String(error),
      );
    } finally {
      setIsImporting(false);
    }
  }

  async function minimizeMainWindow() {
    try {
      const mainWindow = await WebviewWindow.getByLabel("main");
      if (mainWindow) {
        await mainWindow.minimize();
      }
      setClosePromptOpen(false);
    } catch (error) {
      setAppError(
        error instanceof Error ? error.message : String(error),
      );
    }
  }

  async function hideMainWindowToTray() {
    try {
      const mainWindow = await WebviewWindow.getByLabel("main");
      if (mainWindow) {
        await mainWindow.hide();
      }
      setClosePromptOpen(false);
    } catch (error) {
      setAppError(
        error instanceof Error ? error.message : String(error),
      );
    }
  }

  return (
  <main className="app">
    <AppUpdater />

    <header className="app-header">
        <div>
          <h1>PoE Collector</h1>
          <p className="subtitle">
            Path of Exile Unique Collection Tracker
          </p>
        </div>

        <button
          className="settings-button"
          type="button"
          onClick={() => setSettingsOpen(true)}
        >
          Settings
        </button>
      </header>

      <section className="collection-card">
        <div className="collection-info">
          <span className="label">Active Collection</span>

          <select
            className="collection-profile-select"
            value={activeProfileId}
            disabled={!databaseReady}
            onChange={(event) =>
              void handleCollectionProfileChange(
                event.target.value,
              )
            }
          >
            {collectionProfiles.map((profile) => (
              <option
                key={profile.id}
                value={profile.id}
              >
                {profile.name}
              </option>
            ))}
          </select>

          {rolloverPreview &&
  !rolloverPreviewOpen &&
  (rolloverMode === "pending" ||
    rolloverMode === "dev-pending") && (
    <button
      type="button"
      className="catalogue-update-secondary"
      style={{
        alignSelf: "flex-start",
        marginTop: "6px",
      }}
      onClick={() => {
        setRolloverChangesExpanded(false);
        setRolloverPreviewOpen(true);
      }}
    >
      ⚠ {rolloverPreview.oldLeagueName} may have ended
      {" — "}
      Review
    </button>
  )}

          <span className="count">
            {databaseReady
              ? `${uniques.length.toLocaleString()} uniques loaded`
              : "Starting local database..."}
          </span>

          {activeProfile?.kind === "challenge" && (
            <span className="source-file">
              Fresh challenge-league collection
            </span>
          )}

          {activeProfile?.kind === "standard" &&
            sourceFile && (
              <span className="source-file">
                Imported from {sourceFile}
              </span>
            )}
        </div>

          <div className="collection-actions">
  <div className="collection-action-buttons">
    <button
      className="import-button"
      type="button"
      onClick={handleImport}
      disabled={isImporting || snapshotBusy || !databaseReady}
    >
      {isImporting
        ? "Importing..."
        : uniques.length > 0
          ? "Import Another Collection"
          : "Import Existing Collection"}
    </button>

    <button
      className="snapshot-button"
      type="button"
      onClick={() => void handleExportTrackerSnapshot()}
      disabled={snapshotBusy || !databaseReady}
    >
      {snapshotBusy ? "Working..." : "Export Snapshot"}
    </button>

    <button
      className="snapshot-button"
      type="button"
      onClick={() => void handleChooseTrackerSnapshot()}
      disabled={snapshotBusy || !databaseReady}
    >
      Restore Snapshot
    </button>
  </div>

  {snapshotMessage && (
    <span className="snapshot-message">
      {snapshotMessage}
    </span>
  )}
</div>
      </section>

      {appError && (
        <section className="error-message">
          <strong>Something went wrong</strong>
          <span>{appError}</span>
        </section>
      )}

      
      {uniques.length === 0 ? (
        <section className="empty-state">
          {appError ? (
            <>
              <h2>Unable to load collection</h2>
              <p>
                An error occurred while loading your collection.
              </p>
            </>
          ) : !databaseReady ||
            loadedCollectionProfileId !== activeProfileId ||
            catalogueChecking ? (
            <>
              <h2>Loading collection...</h2>
              <p>
                Preparing your {activeProfile?.name ?? "Path of Exile"} collection.
              </p>
            </>
          ) : (
            <>
              <h2>No collection loaded yet</h2>
              <p>
                Import your existing spreadsheet to start tracking your Path of
                Exile uniques.
              </p>
            </>
          )}
        </section>
      ) : (

        <>
          <section className="status-summary">
            {trackingCounts.map(({ label, count }) => (
              <div className="status-card" key={label}>
                <span>{label}</span>
                <strong>{count.toLocaleString()}</strong>
              </div>
            ))}
          </section>

          <section className="filter-panel">
            <div className="search-field">
  <label htmlFor="unique-search">Search</label>

  <div className="search-input-wrap">
    <input
      id="unique-search"
      type="text"
      placeholder="Search uniques..."
      value={searchTerm}
      onChange={(event) =>
        setSearchTerm(event.target.value)
      }
    />

    {searchTerm && (
      <button
        type="button"
        className="search-clear-button"
        aria-label="Clear search"
        title="Clear search"
        onClick={() => setSearchTerm("")}
      >
        ×
      </button>
    )}
  </div>
</div>

            <div className="filter-field">
              <label htmlFor="sort-mode">Sort</label>

              <select
                id="sort-mode"
                value={sortMode}
                onChange={(event) =>
                  handleSortModeChange(event.target.value as SortMode)
                }
              >
                <option value="alphabetical">All A–Z</option>
                <option value="type">By Type → A–Z</option>
              </select>
            </div>

            <div className="filter-field">
              <label htmlFor="status-filter">Status</label>

              <select
                id="status-filter"
                value={statusFilter}
                onChange={(event) =>
                  setStatusFilter(event.target.value as StatusFilter)
                }
              >
                <option value="all">All Statuses</option>
                <option value="unreviewed">Unreviewed</option>
                <option value="missing">Missing</option>
                <option value="owned">Owned</option>
                <option value="wearing">Wearing</option>

                {extraTracking.foil && (
                  <option value="foil">Foil</option>
                )}

                {extraTracking.foulborn && (
                  <option value="foulborn">Foulborn</option>
                )}

                {extraTracking.vestigial && (
                  <option value="vestigial">Vestigial</option>
                )}
              </select>
            </div>

            <div className="filter-field">
              <label htmlFor="type-filter">Type</label>

              <select
                id="type-filter"
                value={typeFilter}
                onChange={(event) => setTypeFilter(event.target.value)}
              >
                <option value="All">All Types</option>

                {availableTypes.map((itemType) => (
                  <option value={itemType} key={itemType}>
                    {itemType}
                  </option>
                ))}
              </select>
            </div>

            <button
              type="button"
              className="catalogue-update-secondary"
              style={{ alignSelf: "flex-end", whiteSpace: "nowrap" }}
              disabled={
                searchTerm === "" &&
                statusFilter === "all" &&
                typeFilter === "All"
              }
              onClick={() => {
                setSearchTerm("");
                setStatusFilter("all");
                setTypeFilter("All");
              }}
            >
              Clear Filters
            </button>
          </section>

          <section className="collection-list">
            <div className="list-header">
              <div>
                <h2>Collection</h2>
                <p>
                  Showing {displayedUniques.length.toLocaleString()} of{" "}
                  {uniques.length.toLocaleString()} uniques.
                </p>
              </div>

              <span>
                {displayedUniques.length.toLocaleString()} shown
              </span>
            </div>

            <div className="unique-table">
              {displayedUniques.length === 0 ? (
                <div className="no-results">
                  No uniques match your current filters.
                </div>
              ) : sortMode === "alphabetical" ? (
                <>
                  <div className="unique-row table-heading">
                    <span>Name</span>
                    <span>Type</span>
                    <span>Tracking</span>
                  </div>

                  {displayedUniques.map((unique) => (
                    <div className="unique-row" key={unique.id}>
                      <span className="unique-name-cell">
                        <span className="unique-name-stack">
                          <span className="unique-name">{unique.name}</span>
                          {(unique.baseType || unique.variantLabel) && (
                            <span className="unique-variant-meta">
                              {[unique.baseType, unique.variantLabel]
                                .filter(Boolean)
                                .join(" • ")}
                            </span>
                          )}
                        </span>
                        {!unique.reviewed && (
                          <span className="unreviewed-badge">
                            NEW / UNREVIEWED
                          </span>
                        )}
                        {unique.isLegacyOnly && (
                          <span className="legacy-badge">
                            LEGACY
                          </span>
                        )}
                      </span>
                      <span className="unique-type">{unique.itemType}</span>

                      <TrackingBadges
                        flags={unique.flags}
                        colors={statusColors}
                        extraTracking={extraTracking}
editionAvailability={
  unique.editionAvailability
}

editionSources={
  unique.editionSources
}
                        isCollected={isUniqueCollected(unique)}
isMissing={isUniqueMissing(unique)}
                        onToggle={(flag) =>
                          toggleTrackingFlag(unique.id, flag)
                        }
                        onMarkMissing={() =>
                          markUniqueMissing(unique.id)
                        }
                      />
                    </div>
                  ))}
                </>
              ) : (
                groupedUniques.map((group) => (
                  <section className="type-group" key={group.itemType}>
                    <div className="type-group-header">
                      <span className="type-group-header-label">
                        {group.itemType.toUpperCase()}
                      </span>
                    </div>

                    {group.items.map((unique) => (
                      <div
                        className="unique-row grouped-row"
                        key={unique.id}
                      >
                        <span className="unique-name-cell">
                          <span className="unique-name-stack">
                            <span className="unique-name">{unique.name}</span>
                            {(unique.baseType || unique.variantLabel) && (
                              <span className="unique-variant-meta">
                                {[unique.baseType, unique.variantLabel]
                                  .filter(Boolean)
                                  .join(" • ")}
                              </span>
                            )}
                          </span>
                          {!unique.reviewed && (
                            <span className="unreviewed-badge">
                              NEW / UNREVIEWED
                            </span>
                          )}
                          {unique.isLegacyOnly && (
                            <span className="legacy-badge">
                              LEGACY
                            </span>
                          )}
                        </span>

                        <TrackingBadges
                          flags={unique.flags}
                          colors={statusColors}
                          extraTracking={extraTracking}
editionAvailability={
  unique.editionAvailability
}

editionSources={
  unique.editionSources
}
                          isCollected={isUniqueCollected(unique)}
isMissing={isUniqueMissing(unique)}
                          onToggle={(flag) =>
                            toggleTrackingFlag(unique.id, flag)
                          }
                          onMarkMissing={() =>
                            markUniqueMissing(unique.id)
                          }
                        />
                      </div>
                    ))}
                  </section>
                ))
              )}
            </div>
                    </section>

          {hasUncertainEditionData && (
            <p className="edition-data-note">
              *Data inconclusive. Starred editions
              may or may not exist.
            </p>
          )}
        </>
      )}

      {rolloverPreviewOpen &&
  rolloverPreview && (
    <div className="catalogue-update-overlay">
      <section className="catalogue-update-modal">
        <div className="catalogue-update-heading">
          <span className="catalogue-update-kicker">
            {rolloverMode === "dev-pending"
              ? "DEV NEW LEAGUE TEST"
              : rolloverMode === "dev-complete"
                ? "DEV ROLLOVER TEST PASSED"
                : rolloverMode === "pending"
                  ? "REVIEW LEAGUE ARCHIVE"
                  : "LEAGUE ARCHIVED"}
          </span>

          <h2>
            {rolloverPreview.oldLeagueName}
            {" → "}
            {rolloverPreview.newLeagueName}
          </h2>

          <p>
            {rolloverMode === "dev-pending"
              ? "Development simulation. The merge runs inside a transaction and is rolled back; no data will change."
              : rolloverMode === "dev-complete"
                ? "The rollover test was rolled back successfully. No collection data changed."
                
: rolloverMode === "pending"
  ? `Review the transfer from ${rolloverPreview.oldLeagueName} into Standard. Only proceed if this is a finished league collection. Archiving adds its collected statuses to Standard without removing existing ownership, preserves the original league's history, and leaves other active leagues untouched. A backup is required before confirming.`

                  : `${rolloverPreview.oldLeagueName} was archived and its owned statuses merged into Standard. Other active leagues were untouched.`}
          </p>
        </div>

        <div className="catalogue-update-stats">
          <div>
            <strong>
              {rolloverPreview.changedUniques}
            </strong>
            <span>uniques changed</span>
          </div>

          <div>
            <strong>
              {rolloverPreview.flagCounts.owned}
            </strong>
            <span>Owned added</span>
          </div>

          <div>
            <strong>
              {rolloverPreview.flagCounts.wearing}
            </strong>
            <span>Wearing added</span>
          </div>

          <div>
            <strong>
              {rolloverPreview.flagCounts.foil}
            </strong>
            <span>Foils added</span>
          </div>
        </div>

        <p className="catalogue-update-note">
          Foulborn added:{" "}
          {rolloverPreview.flagCounts.foulborn}
          {" • "}
          Vestigial added:{" "}
          {rolloverPreview.flagCounts.vestigial}
        </p>

        {rolloverChangesExpanded &&
          rolloverPreview.changes.length > 0 && (
            <div
              style={{
                maxHeight: "220px",
                overflowY: "auto",
              }}
            >
              {rolloverPreview.changes.map(
                (change) => (
                  <p
                    className="catalogue-update-note"
                    key={change.uniqueId}
                  >
                    <strong>
                      {change.name}
                    </strong>
                    {" — "}
                    {change.addedFlags
                      .map(
                        (flag) =>
                          STATUS_LABELS[flag],
                      )
                      .join(", ")}
                  </p>
                ),
              )}
            </div>
          )}

        <div className="catalogue-update-actions">
          {rolloverPreview.changes.length > 0 && (
            <button
              type="button"
              className="catalogue-update-secondary"
              onClick={() =>
                setRolloverChangesExpanded(
                  (current) => !current,
                )
              }
            >
              {rolloverChangesExpanded
                ? "Hide Standard Changes"
                : "View Standard Changes"}
            </button>
          )}

          {(rolloverMode === "pending" ||
            rolloverMode === "dev-pending") && (
            <button
              type="button"
              className="catalogue-update-secondary"
              disabled={rolloverApplying}
              onClick={() =>
                setRolloverPreviewOpen(false)
              }
            >
              Not Now
            </button>
          )}

          {rolloverMode === "pending" ? (
            <button
              type="button"
              className="catalogue-update-primary"
              disabled={rolloverApplying}
              onClick={() =>
                void confirmLeagueRollover()
              }
            >
              {rolloverApplying
                ? "Saving Backup / Archiving..."
                : "Back Up & Archive to Standard"}
            </button>
          ) : rolloverMode ===
            "dev-pending" ? (
            <button
              type="button"
              className="catalogue-update-primary"
              disabled={rolloverApplying}
              onClick={() =>
                void confirmDevLeagueRollover()
              }
            >
              {rolloverApplying
                ? "Testing Rollover..."
                : "Test Archive to Standard"}
            </button>
          ) : (
            
            <button
              type="button"
              className="catalogue-update-primary"
              onClick={() => {
                if (rolloverMode === "complete") {
                  const nextRollover =
                    completedRolloverQueueRef.current.shift();

                  if (nextRollover) {
                    setRolloverPreview(nextRollover);
                    setRolloverChangesExpanded(false);
                    return;
                  }
                }

                setRolloverPreviewOpen(false);
              }}
            >
              {rolloverMode === "dev-complete"
                ? "Close Test"
                : completedRolloverQueueRef.current.length > 0
                  ? "Next"
                  : "Continue"}
            </button>

          )}
        </div>
      </section>
    </div>
  )}


            {newLeagueNotice &&
        !rolloverPreviewOpen && (
          <div className="catalogue-update-overlay">
            <section className="catalogue-update-modal">
              <div className="catalogue-update-heading">
                <span className="catalogue-update-kicker">
                  NEW LEAGUE DETECTED
                </span>

                <h2>{newLeagueNotice}</h2>

                <p>
                  A new Path of Exile league has been detected!
                  A fresh collection has been created for you.

                  All previously known uniques begin as Missing.
                  Your Standard collection and previous league
                  collections remain separate.
                </p>
              </div>

              <div className="catalogue-update-actions">
                <button
                  type="button"
                  className="catalogue-update-primary"
                  onClick={() => setNewLeagueNotice(null)}
                >
                  Start Collecting!
                </button>
              </div>
            </section>
          </div>
        )}

¨¨

      
      {catalogueUpdateOpen &&
        catalogueUpdate &&
        !rolloverPreviewOpen &&
        !newLeagueNotice && (

        <div className="catalogue-update-overlay">
          <section className="catalogue-update-modal">
            <div className="catalogue-update-heading">
              <span className="catalogue-update-kicker">
                CATALOGUE UPDATED
              </span>
              
              <h2>Unique Catalogue Updated</h2>
              <p>
                PoE Collector has downloaded the latest unique item
                information from PoE Wiki. Your existing collection
                progress has been preserved.
              </p>

            </div>

            <div className="catalogue-update-stats">
              <div>
                <strong>{catalogueUpdate.newFamilies}</strong>
                <span>new uniques</span>
              </div>
              <div>
                <strong>{catalogueUpdate.newVariants}</strong>
                <span>new variants</span>
              </div>
              <div>
                <strong>{catalogueUpdate.dropDisabled}</strong>
                <span>drop-disabled</span>
              </div>
              <div>
                <strong>{catalogueUpdate.updatedEntries}</strong>
                <span>updated entries</span>
              </div>
            </div>

            
            <p className="catalogue-update-note">
              Newly discovered uniques start as Missing.
              New variants of existing uniques start as Unreviewed,
              so you can review them without losing your existing
              collection progress.
            </p>


            <div className="catalogue-update-actions">
              <button
                type="button"
                className="catalogue-update-secondary"
                onClick={() => acknowledgeCatalogueUpdate(false)}
              >
                Got it
              </button>
              <button
                type="button"
                className="catalogue-update-primary"
                onClick={() => acknowledgeCatalogueUpdate(true)}
              >
                Review Unreviewed Items
              </button>
            </div>
          </section>
        </div>
      )}

      {pendingImport && (
        <div className="catalogue-update-overlay">
          <section
            className="catalogue-update-modal"
            style={{
              width: "min(760px, calc(100vw - 40px))",
              maxWidth: 760,
            }}
          >
            <div className="catalogue-update-heading">
              <span className="catalogue-update-kicker">
                IMPORT COLLECTION
              </span>
              <h2>How does this spreadsheet track uniques?</h2>
              <p>
                PoE Collector detected a likely format, but you choose how the
                spreadsheet should be interpreted before anything is changed.
              </p>
            </div>

            <div className="settings-toggle-list">
              <label className="settings-toggle-row">
                <span>
                  <strong>Full status list</strong>
                  <br />
                  <small>
                    The sheet lists uniques with statuses such as Owned,
                    Wearing, Missing, or Foil.
                  </small>
                </span>
                <input
                  type="radio"
                  name="import-mode"
                  checked={pendingImportMode === "status-list"}
                  onChange={() =>
                    setPendingImportMode("status-list")
                  }
                />
              </label>

              <label className="settings-toggle-row">
                <span>
                  <strong>Missing-only list</strong>
                  {pendingImport.suggestedMode === "missing-only" && (
                    <>
                      {" "}
                      <span className="legacy-badge">DETECTED</span>
                    </>
                  )}
                  <br />
                  <small>
                    Items still written in the sheet are Missing. Older,
                    single-version catalogue uniques absent from the sheet are
                    treated as Owned.
                  </small>
                </span>
                <input
                  type="radio"
                  name="import-mode"
                  checked={pendingImportMode === "missing-only"}
                  onChange={() =>
                    setPendingImportMode("missing-only")
                  }
                />
              </label>
            </div>

            <div
              style={{
                marginTop: 14,
                display: "grid",
                gap: 6,
              }}
            >
              <label
                className="settings-help"
                htmlFor="spreadsheet-import-destination"
              >
                Save imported collection to
              </label>
              <select
                id="spreadsheet-import-destination"
                className="collection-profile-select"
                value={pendingImportProfileId}
                disabled={isImporting}
                onChange={(event) =>
                  setPendingImportProfileId(
                    event.target.value,
                  )
                }
              >
                {collectionProfiles.map((profile) => (
                  <option
                    key={profile.id}
                    value={profile.id}
                  >
                    {profile.name}
                  </option>
                ))}
              </select>
              <p
                className="catalogue-update-note"
                style={{ margin: 0 }}
              >
                The spreadsheet and any safe missing-list inference will only
                change this collection.
              </p>
            </div>

            {pendingImportMode === "missing-only" && (
              <div
                style={{
                  marginTop: 14,
                  border: "1px solid rgba(255,255,255,0.12)",
                  borderRadius: 10,
                  padding: 14,
                }}
              >
                <h3 style={{ marginTop: 0 }}>
                  New-league safety check
                </h3>

                {pendingImport.latestReleaseLine ? (
                  <>
                    <p className="settings-help">
                      The newest catalogue release is {pendingImport.latestReleaseLine}.
                      Uniques introduced in that release are <strong>not</strong>
                      assumed Owned just because they are absent from this older
                      spreadsheet. PoE Collector leaves their existing collection
                      state alone instead of inferring Owned.
                    </p>

                    <div className="catalogue-update-stats">
                      <div>
                        <strong>{pendingImport.rows.length}</strong>
                        <span>rows in spreadsheet</span>
                      </div>
                      <div>
                        <strong>
                          {pendingImport.protectedLeagueUniques.length}
                        </strong>
                        <span>new-release uniques protected</span>
                      </div>
                    </div>

                    {pendingImport.protectedLeagueUniques.length > 0 && (
                      <>
                        <button
                          type="button"
                          className="catalogue-update-secondary"
                          onClick={() =>
                            setImportDetailsExpanded(
                              (current) => !current,
                            )
                          }
                        >
                          {importDetailsExpanded
                            ? "Hide New Uniques"
                            : "Show New Uniques"}
                        </button>

                        {importDetailsExpanded && (
                          <div
                            style={{
                              maxHeight: 220,
                              overflowY: "auto",
                              marginTop: 10,
                            }}
                          >
                            {pendingImport.protectedLeagueUniques.map(
                              (item) => (
                                <p
                                  className="catalogue-update-note"
                                  key={item.id}
                                >
                                  <strong>{item.name}</strong>
                                  {item.variantLabel
                                    ? ` — ${item.variantLabel}`
                                    : ""}
                                  {" • "}
                                  {item.itemType}
                                </p>
                              ),
                            )}
                          </div>
                        )}
                      </>
                    )}
                  </>
                ) : (
                  <p className="catalogue-update-note">
                    PoE Collector could not determine the newest release from
                    catalogue metadata. For safety, it will not infer any Owned
                    items from absence until that information is available.
                  </p>
                )}

                <p className="catalogue-update-note">
                  Legacy-only entries and families with multiple exact variants
                  are also never assumed Owned from absence alone.
                </p>
              </div>
            )}

            <p className="catalogue-update-note">
              File: <strong>{pendingImport.fileName}</strong>
            </p>

            <div className="catalogue-update-actions">
              <button
                type="button"
                className="catalogue-update-secondary"
                disabled={isImporting}
                onClick={() => {
                  setPendingImport(null);
                  setImportDetailsExpanded(false);
                }}
              >
                Cancel
              </button>
              <button
                type="button"
                className="catalogue-update-primary"
                disabled={isImporting}
                onClick={() => void confirmPendingImport()}
              >
                {isImporting
                  ? "Importing..."
                  : pendingImportMode === "missing-only"
                    ? "Import Missing-Only List"
                    : "Import Status List"}
              </button>
            </div>
          </section>
        </div>
      )}

      {missingOnlySummary && (
        <div className="catalogue-update-overlay">
          <section
            className="catalogue-update-modal"
            style={{
              width: "min(760px, calc(100vw - 40px))",
              maxWidth: 760,
            }}
          >
            <div className="catalogue-update-heading">
              <span className="catalogue-update-kicker">
                MISSING-LIST IMPORT COMPLETE
              </span>
              <h2>Collection inferred safely</h2>
              <p>
                PoE Collector used the missing list to fill {
                  missingOnlySummary.destinationProfileName
                } without pretending that uncertain or newly released uniques
                are already owned.
              </p>
            </div>

            <div className="catalogue-update-stats">
              <div>
                <strong>{missingOnlySummary.explicitMissing}</strong>
                <span>confirmed Missing</span>
              </div>
              <div>
                <strong>{missingOnlySummary.inferredOwned}</strong>
                <span>inferred Owned</span>
              </div>
              <div>
                <strong>
                  {missingOnlySummary.protectedNewLeague.length}
                </strong>
                <span>new-release protected</span>
              </div>
              <div>
                <strong>{missingOnlySummary.protectedAmbiguous}</strong>
                <span>variant/uncertain protected</span>
              </div>
            </div>

            {missingOnlySummary.latestReleaseLine &&
              missingOnlySummary.protectedNewLeague.length > 0 && (
                <>
                  <p className="catalogue-update-note">
                    These uniques are new in release {missingOnlySummary.latestReleaseLine}
                    and were not found in the spreadsheet, so PoE Collector did
                    <strong> not</strong> mark them Owned.
                  </p>

                  <button
                    type="button"
                    className="catalogue-update-secondary"
                    onClick={() =>
                      setImportDetailsExpanded(
                        (current) => !current,
                      )
                    }
                  >
                    {importDetailsExpanded
                      ? "Hide New Uniques"
                      : "Show New Uniques"}
                  </button>

                  {importDetailsExpanded && (
                    <div
                      style={{
                        maxHeight: 240,
                        overflowY: "auto",
                        marginTop: 10,
                      }}
                    >
                      {missingOnlySummary.protectedNewLeague.map(
                        (item) => (
                          <p
                            className="catalogue-update-note"
                            key={item.id}
                          >
                            <strong>{item.name}</strong>
                            {item.variantLabel
                              ? ` — ${item.variantLabel}`
                              : ""}
                            {" • "}
                            {item.itemType}
                          </p>
                        ),
                      )}
                    </div>
                  )}
                </>
              )}

            <div className="catalogue-update-actions">
              {importMatchSummary &&
                (importMatchSummary.ambiguous > 0 ||
                  importMatchSummary.unmatched > 0) && (
                  <button
                    type="button"
                    className="catalogue-update-secondary"
                    onClick={() => {
                      setMissingOnlySummary(null);
                      setImportDetailsExpanded(false);
                      setBatchImportReviewOpen(true);
                    }}
                  >
                    Review Uncertain Imports
                  </button>
                )}

              <button
                type="button"
                className="catalogue-update-primary"
                onClick={() => {
                  setMissingOnlySummary(null);
                  setImportDetailsExpanded(false);
                }}
              >
                Done
              </button>
            </div>
          </section>
        </div>
      )}

      {batchImportReviewOpen && database && (
  <ImportReviewModal
    database={database}
    profiles={collectionProfiles}
    uniques={uniques}
    statusColors={statusColors}
    collectionRules={collectionRules}
    extraTracking={extraTracking}
    initialProfileId={pendingImportProfileId}
    onSaved={async (profileId) => {
      if (profileId === activeProfileId) {
        await loadCollectionData(
          database,
          activeProfileId,
        );
      }
    }}
    onClose={() =>
      setBatchImportReviewOpen(false)
    }
  />
)}

      {pendingSnapshotRestore && (
        <div className="catalogue-update-overlay">
          <section className="catalogue-update-modal">
            <div className="catalogue-update-heading">
              <span className="catalogue-update-kicker">SNAPSHOT RESTORE</span>
              <h2>Restore your collections?</h2>
              <p>
                File: <strong>{pendingSnapshotRestore.fileName}</strong>
              </p>
              <p>
                This will replace all saved league profiles, tracking statuses,
                and reviewed flags with the snapshot data. The current catalogue
                itself will stay in place.
              </p>
              <p>
                {backupMode === "automatic"
                  ? "A backup of your current collection will be created before restoring. Choose a folder if prompted. Restore will stop if the backup fails."
                  : backupMode === "ask"
                    ? "You will be asked whether to back up your current collection before restoring."
                    : "WARNING: Automatic backups are disabled. Restoring will replace your current collection WITHOUT a backup."}
              </p>
              <p>
                {pendingSnapshotRestore.profiles.length} collections ·{" "}
                {pendingSnapshotRestore.collectionRows.toLocaleString()} saved
                catalogue rows
              </p>
              <div style={{ maxHeight: 120, overflowY: "auto" }}>
                {pendingSnapshotRestore.profiles.map((profile) => (
                  <p key={profile.id} className="catalogue-update-note">
                    {profile.name}{profile.isArchived ? " (archived)" : ""}
                  </p>
                ))}
              </div>
              {snapshotRestoreError && (
                <p style={{ color: "#e5b3aa" }}>{snapshotRestoreError}</p>
              )}
            </div>
            <div className="catalogue-update-actions">
              <button
                type="button"
                className="catalogue-update-secondary"
                disabled={snapshotBusy}
                onClick={() => {
                  setPendingSnapshotRestore(null);
                  setSnapshotRestoreError("");
                  setBackupPromptOpen(false);
                }}
              >
                Cancel
              </button>
              {backupMode === "ask" && backupPromptOpen ? (
                <>
                  <button
                    type="button"
                    className="catalogue-update-secondary"
                    disabled={snapshotBusy}
                    onClick={() => void handleConfirmTrackerSnapshotRestore(false)}
                  >
                    Restore Without Backup
                  </button>
                  <button
                    type="button"
                    className="catalogue-update-primary"
                    disabled={snapshotBusy}
                    onClick={() => void handleConfirmTrackerSnapshotRestore(true)}
                  >
                    Back Up & Restore
                  </button>
                </>
              ) : (
                <button
                  type="button"
                  className="catalogue-update-primary"
                  disabled={snapshotBusy}
                  onClick={() => void handleConfirmTrackerSnapshotRestore()}
                >
                  {snapshotBusy
                    ? "Working..."
                    : backupMode === "automatic"
                      ? "Back Up & Restore"
                      : backupMode === "ask"
                        ? "Choose Backup Option"
                        : "Restore Without Backup"}
                </button>
              )}
            </div>
          </section>
        </div>
      )}

      {closePromptOpen && (
        <div className="catalogue-update-overlay">
          <section className="catalogue-update-modal">
            <div className="catalogue-update-heading">
              <span className="catalogue-update-kicker">CLOSE TRACKER</span>
              <h2>What should the tracker do?</h2>
              <p>
                Minimize keeps it on your taskbar. Close to Tray hides the
                window but keeps the in-game lookup hotkey running. To quit
                completely, use Exit Application from the tray icon.
              </p>
            </div>

            <div className="catalogue-update-actions">
              <button
                type="button"
                className="catalogue-update-secondary"
                onClick={() => setClosePromptOpen(false)}
              >
                Cancel
              </button>
              <button
                type="button"
                className="catalogue-update-secondary"
                onClick={() => void minimizeMainWindow()}
              >
                Minimize
              </button>
              <button
                type="button"
                className="catalogue-update-primary"
                onClick={() => void hideMainWindowToTray()}
              >
                Close to Tray
              </button>
            </div>
          </section>
        </div>
      )}

      {settingsOpen && (
        <div
          className="settings-overlay"
          onMouseDown={() => setSettingsOpen(false)}
        >
          <section
            className="settings-modal"
            onMouseDown={(event) => event.stopPropagation()}
          >
            <div className="settings-modal-header">
              <div>
                <h2>Settings</h2>
                <p>Customize how PoE Collector looks and behaves.</p>
                {appVersion && (
                  <p style={{ fontSize: 10, opacity: 0.6 }}>
                    Version: {appVersion}
                  </p>
                )}
              </div>

              <button
                className="settings-close"
                type="button"
                onClick={() => setSettingsOpen(false)}
              >
                ×
              </button>
            </div>

            <div className="settings-section">
              <h3>Backup & Restore</h3>
              <p style={{ color: "#a79d8e", fontSize: 12 }}>
                These preferences apply when you restore a tracker snapshot.
                Normal collection edits do not create separate Excel backups.
              </p>
              <div style={{ display: "flex", flexDirection: "column", gap: 10, marginTop: 12 }}>
                <label htmlFor="backup-mode">Before restoring a snapshot</label>
                <select
                  id="backup-mode"
                  value={backupMode}
                  disabled={!databaseReady || snapshotBusy}
                  onChange={(event) => {
                    void saveBackupMode(event.target.value as BackupMode);
                  }}
                >
                  <option value="automatic">Back up automatically</option>
                  <option value="ask">Ask me every time</option>
                  <option value="off">Do not back up</option>
                </select>
                <button
                  type="button"
                  className="catalogue-update-secondary"
                  style={{ alignSelf: "flex-start" }}
                  disabled={!databaseReady || snapshotBusy}
                  onClick={() => void handleChooseBackupFolder()}
                >
                  Choose Backup Folder
                </button>
                <span style={{ fontSize: 11, color: "#a79d8e", overflowWrap: "anywhere" }}>
                  {backupFolder ? `Backup folder: ${backupFolder}` : "No backup folder selected"}
                </span>
                {backupSettingsMessage && (
                  <span style={{ fontSize: 11, color: "#a79d8e" }}>
                    {backupSettingsMessage}
                  </span>
                )}
              </div>
            </div>

            <div className="settings-section">
              <h3>Status Colors</h3>

              <div className="status-color-list">
                {STATUS_KEYS.map((status) => {
                  const currentColor = statusColors[status];
                  const isRainbow =
                    status === "foil" && currentColor === "rainbow";

                  return (
                    <div className="status-color-row" key={status}>
                      <span>{STATUS_LABELS[status]}</span>

                      <div className="status-color-controls">
                        {status === "foil" && (
                          <button
                            type="button"
                            className={`rainbow-choice ${
                              isRainbow ? "selected" : ""
                            }`}
                            onClick={() =>
                              changeStatusColor("foil", "rainbow")
                            }
                          >
                            Rainbow
                          </button>
                        )}

                        <input
                          key={`${status}-${currentColor}`}
                          type="color"
                          defaultValue={
                            isRainbow ? "#9b51e0" : currentColor
                          }
                          onBlur={(event) =>
                            changeStatusColor(
                              status,
                              event.currentTarget.value,
                            )
                          }
                          aria-label={`Choose ${STATUS_LABELS[status]} color`}
                        />
                      </div>
                    </div>
                  );
                })}
              </div>

              <button
                type="button"
                className="reset-settings-button"
                onClick={resetStatusColors}
              >
                Reset Colors
              </button>
            </div>

            <div className="settings-section settings-divider-section">
  <h3>In-Game Hotkey</h3>

  <p className="settings-help">
    Choose the shortcut used to check the
    Path of Exile item under your mouse.
  </p>

  <div
    className={`hotkey-test-status ${
      hotkeyReady ? "ready" : ""
    }`}
  >
    <strong>
      Current hotkey
    </strong>

    <span>
      {formatHotkeyForDisplay(
        lookupHotkey,
      )}
    </span>
  </div>

  {hotkeyRecording ? (
    <button
      autoFocus
      type="button"
      className="reset-settings-button"
      onKeyDown={
        captureLookupHotkey
      }
      onBlur={() =>
        setHotkeyRecording(false)
      }
    >
      Press your new shortcut...
    </button>
  ) : (
    <div className="parser-test-actions">
      <button
        type="button"
        className="reset-settings-button"
        disabled={!database}
        onClick={() => {
          setHotkeySettingMessage("");
          setHotkeyRecording(true);
        }}
      >
        Change Hotkey
      </button>

      <button
        type="button"
        className="reset-settings-button"
        disabled={
          !database ||
          lookupHotkey ===
            DEFAULT_POE_LOOKUP_HOTKEY
        }
        onClick={() =>
          void changeLookupHotkey(
            DEFAULT_POE_LOOKUP_HOTKEY,
          )
        }
      >
        Reset to Default
      </button>
    </div>
  )}

  {hotkeyRecording && (
    <p className="settings-help">
      Press the shortcut now.
      Press Escape to cancel.
    </p>
  )}

  {hotkeySettingMessage && (
    <p className="catalogue-check-status">
      {hotkeySettingMessage}
    </p>
  )}
</div>

            <div className="settings-section settings-divider-section">
              <h3>What Counts as Collected?</h3>

              <p className="settings-help">
                Choose which kinds of copies prevent a unique from being
                considered Missing.
              </p>

              <div className="settings-toggle-list">
                <label className="settings-toggle-row">
                  <span>Owned</span>

                  <input
                    type="checkbox"
                    checked={collectionRules.owned}
                    onChange={(event) =>
                      changeCollectionRule("owned", event.target.checked)
                    }
                  />
                </label>

                <label className="settings-toggle-row">
                  <span>Wearing</span>

                  <input
                    type="checkbox"
                    checked={collectionRules.wearing}
                    onChange={(event) =>
                      changeCollectionRule("wearing", event.target.checked)
                    }
                  />
                </label>

                {EXTRA_VARIANTS.map((variant) => (
                  <label className="settings-toggle-row" key={variant}>
                    <span>{STATUS_LABELS[variant]}</span>

                    <input
                      type="checkbox"
                      checked={collectionRules[variant]}
                      onChange={(event) =>
                        changeCollectionRule(
                          variant,
                          event.target.checked,
                        )
                      }
                    />
                  </label>
                ))}
              </div>
            </div>

            <div className="settings-section settings-divider-section">
              <h3>Extra Variants to Track</h3>

              <p className="settings-help">
                Hidden variants keep their saved collection data and can be
                enabled again later.
              </p>

              <div className="settings-toggle-list">
                {EXTRA_VARIANTS.map((variant) => (
                  <label className="settings-toggle-row" key={variant}>
                    <span>{STATUS_LABELS[variant]}</span>

                    <input
                      type="checkbox"
                      checked={extraTracking[variant]}
                      onChange={(event) =>
                        changeExtraTracking(
                          variant,
                          event.target.checked,
                        )
                      }
                    />
                  </label>
                ))}
              </div>
            </div>

            <div className="settings-section settings-divider-section">
              <h3>Catalogue Display</h3>

              <p className="settings-help">
                Legacy variants remain saved in your collection even when
                hidden.
              </p>

              <div className="settings-toggle-list">
                <label className="settings-toggle-row">
                  <span>Show Legacy Variants</span>

                  <input
                    type="checkbox"
                    checked={showLegacyVariants}
                    onChange={(event) =>
                      changeShowLegacyVariants(
                        event.target.checked,
                      )
                    }
                  />
                </label>
              </div>
            </div>

            <div className="settings-section settings-divider-section">
              <h3>Catalogue Updates</h3>

              <p className="settings-help">
                PoE Collector checks the PoE Wiki catalogue in the background
                at most once every 24 hours. Your local collection still works
                normally while offline.
              </p>

              <p className="catalogue-check-status">
                {catalogueCheckMessage}
              </p>

              <button
                type="button"
                className="reset-settings-button"
                disabled={catalogueChecking || !database}
                onClick={() => {
                  if (database) {
                    void runCatalogueCheck(database, true);
                  }
                }}
              >
                {catalogueChecking
                  ? "Checking Catalogue..."
                  : "Check for Catalogue Updates"}
              </button>
            </div>

            <div className="settings-section settings-divider-section">
              <h3>Imported Collection Matching</h3>

              <p className="settings-help">
                Imported spreadsheets are reference data only. PoE Collector
                matches their statuses onto the canonical Path of Exile
                catalogue instead of creating catalogue entries from imported
                names.
              </p>

              {importMatchSummary ? (
                <div className="import-match-summary">
                  <div>
                    <strong>{importMatchSummary.matched}</strong>
                    <span>matched</span>
                  </div>
                  <div>
                    <strong>{importMatchSummary.ambiguous}</strong>
                    <span>need review</span>
                  </div>
                  <div>
                    <strong>{importMatchSummary.unmatched}</strong>
                    <span>unmatched</span>
                  </div>
                </div>
              ) : (
                <p className="catalogue-check-status">
                  Import matching will run after the canonical catalogue is
                  available.
                </p>
              )}

                            {importMatchSummary &&
                importMatchSummary.skipped > 0 && (
                  <p className="settings-help import-match-note">
                    {importMatchSummary.skipped} imported{" "}
                    {importMatchSummary.skipped === 1
                      ? "row is"
                      : "rows are"}{" "}
                    being ignored by manual choice.
                  </p>
                )}

              {importMatchSummary &&
                (importMatchSummary.ambiguous > 0 ||
                  importMatchSummary.unmatched > 0) && (
                  <>
                    <p className="settings-help import-match-note">
                      Nothing was guessed for uncertain rows.
                      Review them manually to connect the spreadsheet
                      entry to the correct canonical unique.
                    </p>

                    <button
                      type="button"
                      className="reset-settings-button"
                      disabled={!database}
                      onClick={openImportReviewScreen}
                    >
                      Review Matches
                    </button>
                  </>
                )}
            </div>

            {import.meta.env.DEV && (
              <>
            <div className="settings-section settings-divider-section">
              <h3>Item Parser Test</h3>

              <p className="settings-help">
                Temporary development tool: hover a unique in Path of Exile,
                press Ctrl+C, paste the copied item text here, and see whether
                PoE Collector identifies the exact catalogue variant.
              </p>

              <div
                className={`hotkey-test-status ${
                  hotkeyReady ? "ready" : ""
                }`}
              >
                <strong>
                  In-game hotkey:{" "}
{formatHotkeyForDisplay(
  lookupHotkey,
)}
                </strong>
                <span>
                  {hotkeyBusy
                    ? "Looking up hovered item..."
                    : hotkeyMessage}
                </span>
              </div>

              <textarea
                className="parser-test-textarea"
                value={parserTestText}
                onChange={(event) => {
                  setParserTestText(event.target.value);
                  setParserTestResult(null);
                }}
                placeholder={`Item Class: Boots
Rarity: Unique
Ralakesh's Impatience
Riveted Boots
--------
...`}
              />

              <div className="parser-test-actions">
                <button
                  type="button"
                  className="reset-settings-button"
                  disabled={parserTesting || !database}
                  onClick={() =>
                    void readClipboardAndIdentify()
                  }
                >
                  {parserTesting
                    ? "Reading Clipboard..."
                    : "Read Clipboard & Identify"}
                </button>

                <button
                  type="button"
                  className="reset-settings-button"
                  disabled={
                    parserTesting ||
                    !database ||
                    !parserTestText.trim()
                  }
                  onClick={() => void runParserTest()}
                >
                  {parserTesting
                    ? "Identifying Item..."
                    : "Identify Pasted Item"}
                </button>
              </div>

              {parserTestResult && (
                <div
                  className={`parser-test-result parser-${parserTestResult.status}`}
                >
                  {parserTestResult.status === "matched" ? (
                    <>
                      <strong>Match found</strong>
                      <span>{parserTestResult.name}</span>
                      <small>
                        {[
                          parserTestResult.baseType,
                          parserTestResult.variantLabel,
                        ]
                          .filter(Boolean)
                          .join(" • ") || parserTestResult.itemType}
                      </small>
                      <small>
                        Edition:{" "}
                        {parserTestResult.edition === "foulborn"
                          ? "Foulborn"
                          : parserTestResult.edition === "vestigial"
                            ? "Vestigial"
                            : "Normal"}
                      </small>

                      <small>
                        Matched by{" "}
                        {parserTestResult.matchedBy === "special-rule"
                          ? "special variant rule"
                          : parserTestResult.matchedBy === "variant-label"
                            ? "variant evidence"
                            : "name + base type"}
                      </small>

                      {(() => {
                        const localEntry = uniques.find(
                          (unique) =>
                            unique.id === parserTestResult.uniqueId,
                        );

                        if (!localEntry) {
                          return null;
                        }

                        const editionFlag =
                          parserTestResult.edition === "foulborn"
                            ? "foulborn"
                            : parserTestResult.edition === "vestigial"
                              ? "vestigial"
                              : null;

                        if (!localEntry.reviewed) {
                          return (
                            <small>
                              This edition: Unreviewed
                            </small>
                          );
                        }

                        if (editionFlag) {
                          return (
                            <>
                              <small>
                                This edition:{" "}
                                {localEntry.flags.includes(
                                  editionFlag,
                                )
                                  ? "Owned"
                                  : "Missing"}
                              </small>
                              <small>
                                All tracking:{" "}
                                {localEntry.flags.length > 0
                                  ? localEntry.flags.join(", ")
                                  : "none"}
                              </small>
                            </>
                          );
                        }

                        return (
                          <small>
                            Collection:{" "}
                            {localEntry.flags.length > 0
                              ? localEntry.flags.join(", ")
                              : "Missing"}
                          </small>
                        );
                      })()}
                    </>
                  ) : (
                    <>
                      <strong>
                        {parserTestResult.status === "ambiguous"
                          ? "Ambiguous match"
                          : parserTestResult.status === "not-found"
                            ? "Not found"
                            : parserTestResult.status === "not-unique"
                              ? "Not a unique"
                              : "Could not parse item"}
                      </strong>

                      <span>{parserTestResult.message}</span>

                      {"candidates" in parserTestResult &&
                        parserTestResult.candidates.length > 0 && (
                          <small>
                            Candidates:{" "}
                            {parserTestResult.candidates
                              .map((candidate) =>
                                [
                                  candidate.name,
                                  candidate.baseType,
                                  candidate.variantLabel,
                                ]
                                  .filter(Boolean)
                                  .join(" • "),
                              )
                              .join(" | ")}
                          </small>
                        )}
                    </>
                  )}
                </div>
              )}
            </div>
              </>
            )}
          </section>
        </div>
      )}
    </main>
  );
}

function App() {
  const params = new URLSearchParams(
    window.location.search,
  );

  if (params.get("overlay") === "1") {
    return <OverlayApp />;
  }

  return <MainApp />;
}

export default App;