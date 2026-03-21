(function () {
  const TECH_COUNT = 8;
  /** manhã (4) + extra manhã (1) + almoço (1) + tarde (4) + extra tarde (1) = 11 linhas. */
  const SLOTS_PER_DAY = 11;
  /** Slots fixos em armazenamento (9–13h e 14–18h). Índices 0–7. */
  const BOOKABLE_STORAGE_SLOTS = 8;
  /** Linha «Extra — manhã» (antes do almoço). Armazenamento 8. */
  const MORNING_EXTRA_ROW = 4;
  /** Pausa de almoço — sem marcações. */
  const LUNCH_ROW_INDEX = 5;
  /** Linha «Extra — tarde». Armazenamento 9. */
  const AFTERNOON_EXTRA_ROW = 10;
  const START_HOUR = 9;

  const STORAGE_KEY = "agenda-servicos-v1";
  const CONFIG_STORAGE_KEY = "agenda-servicos-config-v1";
  const GEO_CACHE_PREFIX = "agenda-cp-geo:";

  function defaultAgendaConfig() {
    return {
      names: Array(TECH_COUNT).fill(""),
      baseCps: Array(TECH_COUNT).fill(""),
      googleMapsApiKey: "",
    };
  }

  /** @type {{ names: string[], baseCps: string[], googleMapsApiKey: string }} */
  let __configSnapshot = defaultAgendaConfig();
  let persistConfigTimer = null;
  /** Cache em memória para pares CP→CP (distância km). */
  const postalPairDistanceCache = new Map();
  /** Carregamento único do script Maps JS (REST Geocoding/Distance Matrix não funcionam no browser por CORS). */
  let googleMapsScriptPromise = null;
  /** Chave com que o script `maps/api/js` foi carregado (mudar chave no modal exige F5). */
  let googleMapsKeyUsedForLoad = "";

  function googleMapsApiKey() {
    return (loadConfig().googleMapsApiKey || "").trim();
  }

  /**
   * Garante `google.maps` (Geocoder + DistanceMatrixService). Usa a mesma chave que o modal Técnicos.
   * No Google Cloud: activar **Maps JavaScript API** (e facturação).
   */
  function ensureGoogleMapsLoaded() {
    const gkey = googleMapsApiKey();
    if (!gkey) return Promise.reject(new Error("no_google_key"));
    try {
      if (globalThis.google?.maps?.Geocoder && googleMapsKeyUsedForLoad === gkey) {
        return Promise.resolve();
      }
      if (globalThis.google?.maps?.Geocoder && googleMapsKeyUsedForLoad && googleMapsKeyUsedForLoad !== gkey) {
        console.warn(
          "Chave Google Maps alterada: recarrega a página (F5) para usar a nova chave."
        );
        return Promise.reject(new Error("google_key_changed_reload"));
      }
    } catch {
      /* ignore */
    }
    if (googleMapsScriptPromise) return googleMapsScriptPromise;
    googleMapsScriptPromise = new Promise((resolve, reject) => {
      const cbName = "__agendaGmapsInit_" + String(Date.now());
      globalThis[cbName] = () => {
        try {
          delete globalThis[cbName];
          if (globalThis.google?.maps?.Geocoder) {
            googleMapsKeyUsedForLoad = gkey;
            resolve();
          } else {
            googleMapsScriptPromise = null;
            reject(new Error("maps_not_ready"));
          }
        } catch (e) {
          googleMapsScriptPromise = null;
          reject(e);
        }
      };
      const s = document.createElement("script");
      s.async = true;
      s.defer = true;
      s.src = `https://maps.googleapis.com/maps/api/js?key=${encodeURIComponent(gkey)}&callback=${cbName}`;
      s.onerror = () => {
        try {
          delete globalThis[cbName];
        } catch {
          /* ignore */
        }
        googleMapsScriptPromise = null;
        reject(new Error("maps_script_failed"));
      };
      document.head.appendChild(s);
    });
    return googleMapsScriptPromise;
  }

  /** Raio em km para considerar "mesma zona" entre dois códigos postais. */
  const ZONE_RADIUS_KM = 5;
  /** Acima desta distância (km) ao técnico mais próximo, sugere-se nova rota noutro dia (slot extra). */
  const NEW_ROUTE_DISTANCE_KM = 10;

  const WEEKDAY_SHORT = ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"];

  /** Cliente Supabase (preenchido em startAgenda). */
  let supabaseClient = null;
  /** Canais Supabase Realtime (agenda); removidos no logout. */
  let realtimeAgendaChannels = [];
  let realtimeRenderTimer = null;
  /** Fallback se o WebSocket não entregar eventos (rede / proxy). */
  let agendaPollIntervalId = null;
  /** Espelho em memória das marcações (antes localStorage). */
  let __bookingStore = { __schema: 2 };
  let persistBookingsTimer = null;
  /** Evita que o poll `hydrateBookingsFromCloud` apague marcações ainda a sincronizar. */
  let persistBookingsInFlight = 0;

  function sleep(ms) {
    return new Promise((r) => setTimeout(r, ms));
  }

  /** Evita espera infinita se o Supabase/rede não responder. */
  function promiseWithTimeout(promise, ms, timeoutMessage) {
    return new Promise((resolve, reject) => {
      const t = setTimeout(() => reject(new Error(timeoutMessage)), ms);
      promise.then(
        (v) => {
          clearTimeout(t);
          resolve(v);
        },
        (e) => {
          clearTimeout(t);
          reject(e);
        }
      );
    });
  }

  /** Mensagens no login: azul (estado) vs vermelho (erro). */
  function setAuthFeedback(el, mode, text) {
    if (!el) return;
    if (mode === "hidden") {
      el.hidden = true;
      el.className = "auth-screen__feedback";
      el.removeAttribute("role");
      return;
    }
    el.hidden = false;
    el.textContent = text;
    el.className =
      "auth-screen__feedback " +
      (mode === "loading" ? "auth-screen__feedback--loading" : "auth-screen__feedback--error");
    el.setAttribute("role", mode === "loading" ? "status" : "alert");
  }

  function haversineKm(lat1, lon1, lat2, lon2) {
    const R = 6371;
    const toRad = (d) => (d * Math.PI) / 180;
    const dLat = toRad(lat2 - lat1);
    const dLon = toRad(lon2 - lon1);
    const a =
      Math.sin(dLat / 2) ** 2 +
      Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) ** 2;
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    return R * c;
  }

  function geoCacheKey(normalized) {
    return GEO_CACHE_PREFIX + (googleMapsApiKey() ? "g:" : "osm:") + normalized;
  }

  /** @returns {Promise<{ lat: number, lon: number } | null>} */
  async function geocodePostalPt(cp) {
    const normalized = String(cp || "").trim();
    if (!/^\d{4}-\d{3}$/.test(normalized)) return null;
    const cacheKey = geoCacheKey(normalized);
    try {
      const raw = localStorage.getItem(cacheKey);
      if (raw) {
        const o = JSON.parse(raw);
        if (o && Number.isFinite(o.lat) && Number.isFinite(o.lon) && o.ts && Date.now() - o.ts < 30 * 24 * 60 * 60 * 1000) {
          return { lat: o.lat, lon: o.lon };
        }
      }
    } catch {}

    const gkey = googleMapsApiKey();
    if (gkey) {
      try {
        await promiseWithTimeout(
          ensureGoogleMapsLoaded(),
          15000,
          "Timeout ao carregar Google Maps (ver chave, Maps JavaScript API e referrers)."
        );
        const g = globalThis.google.maps;
        const coords = await promiseWithTimeout(
          new Promise((resolve, reject) => {
            const geocoder = new g.Geocoder();
            geocoder.geocode(
              { address: `${normalized}, Portugal`, region: "pt" },
              (results, status) => {
                if (status === "OK" && results && results[0]) {
                  const loc = results[0].geometry.location;
                  resolve({ lat: loc.lat(), lon: loc.lng() });
                } else {
                  reject(new Error(String(status)));
                }
              }
            );
          }),
          12000,
          "Timeout Geocoder Google."
        );
        if (coords && Number.isFinite(coords.lat) && Number.isFinite(coords.lon)) {
          try {
            localStorage.setItem(cacheKey, JSON.stringify({ lat: coords.lat, lon: coords.lon, ts: Date.now() }));
          } catch {
            /* ignore */
          }
          return coords;
        }
      } catch (e) {
        console.warn("Geocoding Google (Maps JS):", e?.message || e);
      }
    }

    const q = encodeURIComponent(`${normalized}, Portugal`);
    const url = `https://nominatim.openstreetmap.org/search?format=json&limit=1&q=${q}`;
    try {
      await sleep(1100);
      const res = await fetch(url, { headers: { Accept: "application/json" } });
      if (!res.ok) return null;
      const data = await res.json();
      if (!Array.isArray(data) || !data[0]) return null;
      const lat = parseFloat(data[0].lat);
      const lon = parseFloat(data[0].lon);
      if (!Number.isFinite(lat) || !Number.isFinite(lon)) return null;
      try {
        localStorage.setItem(cacheKey, JSON.stringify({ lat, lon, ts: Date.now() }));
      } catch {}
      return { lat, lon };
    } catch {
      return null;
    }
  }

  /**
   * Distância entre dois CP (km). Com chave Google: Distance Matrix (condução), alinhado ao Maps.
   * Sem chave: linha reta após geocodificação OSM.
   * @returns {Promise<number | null>}
   */
  async function postalDistanceKm(cpA, cpB) {
    const a = String(cpA || "").trim();
    const b = String(cpB || "").trim();
    if (!/^\d{4}-\d{3}$/.test(a) || !/^\d{4}-\d{3}$/.test(b)) return null;
    if (a === b) return 0;
    const pairKey = [a, b].sort().join("|");
    if (postalPairDistanceCache.has(pairKey)) return postalPairDistanceCache.get(pairKey);

    const gkey = googleMapsApiKey();
    if (gkey) {
      try {
        await promiseWithTimeout(
          ensureGoogleMapsLoaded(),
          15000,
          "Timeout ao carregar Google Maps."
        );
        const g = globalThis.google.maps;
        const km = await promiseWithTimeout(
          new Promise((resolve, reject) => {
            const service = new g.DistanceMatrixService();
            service.getDistanceMatrix(
              {
                origins: [`${a}, Portugal`],
                destinations: [`${b}, Portugal`],
                travelMode: g.TravelMode.DRIVING,
                unitSystem: g.UnitSystem.METRIC,
              },
              (response, status) => {
                if (
                  status === "OK" &&
                  response?.rows?.[0]?.elements?.[0]?.status === "OK"
                ) {
                  resolve(response.rows[0].elements[0].distance.value / 1000);
                } else {
                  reject(new Error(String(status)));
                }
              }
            );
          }),
          12000,
          "Timeout Distance Matrix Google."
        );
        if (Number.isFinite(km)) {
          postalPairDistanceCache.set(pairKey, km);
          return km;
        }
      } catch (e) {
        console.warn("Distance Matrix Google (Maps JS):", e?.message || e);
      }
    }

    const c1 = await geocodePostalPt(a);
    const c2 = await geocodePostalPt(b);
    if (!c1 || !c2) return null;
    const km = haversineKm(c1.lat, c1.lon, c2.lat, c2.lon);
    postalPairDistanceCache.set(pairKey, km);
    return km;
  }

  /**
   * Outros técnicos com marcação na mesma zona (CP a ≤ ZONE_RADIUS_KM km).
   * @param {string} postalCodeNew — CP do novo serviço (distância via {@link postalDistanceKm})
   * @returns {Promise<string[]>} nomes únicos, ordenados
   */
  async function analyzeNearbyTechniciansZone(bookingDateIso, currentTechIndex, excludeRangeId, postalCodeNew) {
    /** @type {Map<string, string>} dedupe -> cp */
    const uniqueCps = new Map();
    const all = loadAll();

    for (const key of Object.keys(all)) {
      const parts = key.split("|");
      if (parts.length !== 3) continue;
      const dateStr = parts[0];
      const techIndex = Number(parts[1]);
      const b = normalizeBooking(all[key]);
      if (!b || !b.postalCode || !isValidPostalPt(b.postalCode)) continue;

      if (techIndex === currentTechIndex) {
        if (excludeRangeId && b.rangeId === excludeRangeId) continue;
        continue;
      }

      if (dateStr < bookingDateIso) continue;

      const rid = b.rangeId || key;
      const dedupe = `${dateStr}|${techIndex}|${rid}`;
      if (!uniqueCps.has(dedupe)) uniqueCps.set(dedupe, b.postalCode);
    }

    /** @type {Map<string, Set<string>>} nome do técnico -> datas ISO */
    const techToDateSet = new Map();

    for (const [dedupe, cpOther] of uniqueCps) {
      const km = await postalDistanceKm(postalCodeNew, cpOther);
      if (km === null || km > ZONE_RADIUS_KM) continue;

      const dateStr = dedupe.split("|")[0];
      const techIndex = Number(dedupe.split("|")[1]);
      const name = getTechDisplayName(techIndex);
      if (!techToDateSet.has(name)) techToDateSet.set(name, new Set());
      techToDateSet.get(name).add(dateStr);
    }

    const techNames = [...techToDateSet.keys()].sort((a, b) => a.localeCompare(b, "pt"));
    /** @type {Map<string, string[]>} */
    const techToDates = new Map();
    for (const [name, set] of techToDateSet) {
      techToDates.set(name, [...set].sort());
    }

    return { techNames, techToDates };
  }

  /**
   * Sobreposição de zona no mês: CPs a ≤ ZONE_RADIUS_KM no mesmo dia entre técnicos (marcações distintas).
   * @param {string} monthPrefix — "YYYY-MM"
   * @returns {Promise<{ hasOverlap: boolean, conflictingKeys: Set<string> }>}
   * Chave: `${dateStr}|${techIndex}|${rid}` com rid = rangeId da marcação ou chave de armazenamento `date|tech|slot`.
   */
  async function computeSameDayZoneOverlapDetails(monthPrefix) {
    const empty = () => ({ hasOverlap: false, conflictingKeys: new Set() });
    if (!monthPrefix) return empty();
    const all = loadAll();
    /** dateStr -> Map<`${techIndex}|${rid}`, { techIndex: number, cp: string, rid: string }> */
    const dateToEntries = new Map();

    for (const key of Object.keys(all)) {
      if (key === "__schema") continue;
      const parts = key.split("|");
      if (parts.length !== 3) continue;
      const dateStr = parts[0];
      if (!dateStr.startsWith(monthPrefix)) continue;
      const techIndex = Number(parts[1]);
      const b = normalizeBooking(all[key]);
      if (!b || !isValidPostalPt(b.postalCode)) continue;
      const rid = b.rangeId ? String(b.rangeId) : key;
      const mapKey = `${techIndex}|${rid}`;
      if (!dateToEntries.has(dateStr)) dateToEntries.set(dateStr, new Map());
      const m = dateToEntries.get(dateStr);
      if (!m.has(mapKey)) m.set(mapKey, { techIndex, cp: b.postalCode.trim(), rid });
    }

    let hasOverlap = false;
    const conflictingKeys = new Set();

    for (const [dateStr, m] of dateToEntries) {
      const entries = [...m.values()];
      for (let i = 0; i < entries.length; i++) {
        for (let j = i + 1; j < entries.length; j++) {
          const a = entries[i];
          const b = entries[j];
          if (a.techIndex === b.techIndex) continue;
          const km = await postalDistanceKm(a.cp, b.cp);
          if (km === null || km > ZONE_RADIUS_KM) continue;
          hasOverlap = true;
          conflictingKeys.add(`${dateStr}|${a.techIndex}|${a.rid}`);
          conflictingKeys.add(`${dateStr}|${b.techIndex}|${b.rid}`);
        }
      }
    }

    return { hasOverlap, conflictingKeys };
  }

  async function computeSameDayZoneOverlapInMonth(monthPrefix) {
    const d = await computeSameDayZoneOverlapDetails(monthPrefix);
    return d.hasOverlap;
  }

  let zoneOverlapRefreshToken = 0;

  function clearAllSlotZoneDots() {
    document.querySelectorAll(".slot__zone-dot").forEach((d) => d.remove());
  }

  /** @param {Set<string>} conflictingKeys */
  function applySlotZoneConflictDots(conflictingKeys) {
    document.querySelectorAll(".slot.slot--busy[data-range-key]").forEach((btn) => {
      if (btn.classList.contains("slot--busy-span-continuation")) {
        const orphan = btn.querySelector(".slot__zone-dot");
        if (orphan) orphan.remove();
        return;
      }
      const k = btn.getAttribute("data-range-key");
      const want = Boolean(k && conflictingKeys.has(k));
      let dot = btn.querySelector(".slot__zone-dot");
      if (want && !dot) {
        dot = document.createElement("span");
        dot.className = "slot__zone-dot";
        dot.setAttribute("aria-hidden", "true");
        dot.title = "Conflito de zona: CP próximo de outro técnico no mesmo dia";
        btn.appendChild(dot);
      } else if (!want && dot) {
        dot.remove();
      }
    });
  }

  function scheduleZoneOverlapUiRefresh() {
    const el = document.getElementById("zoneOverlapIndicator");
    const monthPrefix = monthInput && monthInput.value ? monthInput.value : "";
    if (!monthPrefix) {
      if (el) el.setAttribute("hidden", "");
      clearAllSlotZoneDots();
      return;
    }
    const token = ++zoneOverlapRefreshToken;
    void (async () => {
      await sleep(320);
      if (token !== zoneOverlapRefreshToken) return;
      try {
        const { hasOverlap, conflictingKeys } = await computeSameDayZoneOverlapDetails(monthPrefix);
        if (token !== zoneOverlapRefreshToken) return;
        if (el) {
          if (hasOverlap) el.removeAttribute("hidden");
          else el.setAttribute("hidden", "");
        }
        applySlotZoneConflictDots(conflictingKeys);
      } catch {
        if (token !== zoneOverlapRefreshToken) return;
        if (el) el.setAttribute("hidden", "");
        clearAllSlotZoneDots();
      }
    })();
  }

  function formatDateLongPt(iso) {
    return new Date(iso + "T12:00:00").toLocaleDateString("pt-PT", {
      weekday: "long",
      day: "numeric",
      month: "long",
      year: "numeric",
    });
  }

  function formatDatesListPt(isoDates) {
    const parts = isoDates.map(formatDateLongPt);
    if (parts.length === 0) return "—";
    if (parts.length === 1) return parts[0];
    if (parts.length === 2) return `${parts[0]} e ${parts[1]}`;
    return `${parts.slice(0, -1).join("; ")} e ${parts[parts.length - 1]}`;
  }

  /**
   * @param {string[]} names
   * @param {Map<string, string[]>} techToDates
   */
  function formatTechZoneMessage(names, techToDates) {
    const n = names.filter(Boolean);
    if (n.length === 0) return "";

    if (n.length === 1) {
      const nm = n[0];
      const dates = techToDates.get(nm) || [];
      const datesStr = formatDatesListPt(dates);
      return `O técnico "${nm}" já está nessa zona.\n\nDias: ${datesStr}.`;
    }

    const head =
      n.length === 2
        ? `Os técnicos "${n[0]}" e "${n[1]}" já estão nessa zona.`
        : `Os técnicos ${n
            .slice(0, -1)
            .map((x) => `"${x}"`)
            .join(", ")} e "${n[n.length - 1]}" já estão nessa zona.`;

    const lines = n.map((nm) => {
      const datesStr = formatDatesListPt(techToDates.get(nm) || []);
      return `• ${nm}: ${datesStr}`;
    });

    return `${head}\n\n${lines.join("\n")}.`;
  }

  /**
   * @returns {Promise<'ok' | 'keep' | 'change' | 'cancel'>}
   */
  async function alertNearbyZoneIfNeeded(postalCode, bookingDateIso, currentTechIndex, excludeRangeId) {
    if (!isValidPostalPt(postalCode)) return "ok";
    const coordsNew = await geocodePostalPt(postalCode);
    if (!coordsNew) {
      window.alert(
        "Não foi possível obter coordenadas deste código postal (rede ou serviço indisponível). Pode guardar a marcação na mesma."
      );
      return "ok";
    }
    const { techNames, techToDates } = await analyzeNearbyTechniciansZone(
      bookingDateIso,
      currentTechIndex,
      excludeRangeId,
      postalCode
    );
    if (techNames.length === 0) return "ok";

    const msg = formatTechZoneMessage(techNames, techToDates);
    return showZoneWarningModal(msg);
  }

  /**
   * @returns {Promise<'keep' | 'change' | 'cancel'>}
   */
  function showZoneWarningModal(message) {
    const dialog = document.getElementById("zoneWarningDialog");
    const textEl = document.getElementById("zoneWarningText");
    const btnKeep = document.getElementById("btnZoneWarningKeep");
    const btnChange = document.getElementById("btnZoneWarningChange");
    if (!dialog || !textEl || !btnKeep || !btnChange) {
      window.alert(`${message}\n\nDeseja manter (OK) ou cancelar e alterar depois?`);
      return Promise.resolve("keep");
    }
    textEl.textContent = message;
    return new Promise((resolve) => {
      const cleanup = () => {
        btnKeep.removeEventListener("click", onKeep);
        btnChange.removeEventListener("click", onChange);
        dialog.removeEventListener("cancel", onEsc);
      };
      const onKeep = () => {
        cleanup();
        dialog.close();
        resolve("keep");
      };
      const onChange = () => {
        cleanup();
        dialog.close();
        resolve("change");
      };
      const onEsc = () => {
        cleanup();
        resolve("cancel");
      };
      btnKeep.addEventListener("click", onKeep);
      btnChange.addEventListener("click", onChange);
      dialog.addEventListener("cancel", onEsc);
      dialog.showModal();
    });
  }

  /** @param {unknown} parsed */
  function parseConfigFromRaw(parsed) {
    if (!parsed || typeof parsed !== "object") return defaultAgendaConfig();
    const o = /** @type {{ names?: unknown, baseCps?: unknown, googleMapsApiKey?: unknown }} */ (parsed);
    const names = Array.isArray(o.names) ? o.names.map(String) : [];
    while (names.length < TECH_COUNT) names.push("");
    const baseCps = Array.isArray(o.baseCps) ? o.baseCps.map(String) : [];
    while (baseCps.length < TECH_COUNT) baseCps.push("");
    const googleMapsApiKey = typeof o.googleMapsApiKey === "string" ? o.googleMapsApiKey : "";
    return {
      names: names.slice(0, TECH_COUNT),
      baseCps: baseCps.slice(0, TECH_COUNT),
      googleMapsApiKey,
    };
  }

  function loadConfigFromLocalStorageIntoSnapshot() {
    try {
      const raw = localStorage.getItem(CONFIG_STORAGE_KEY);
      if (!raw) {
        __configSnapshot = defaultAgendaConfig();
        return;
      }
      __configSnapshot = parseConfigFromRaw(JSON.parse(raw));
    } catch {
      __configSnapshot = defaultAgendaConfig();
    }
  }

  /** @returns {{ names: string[], baseCps: string[], googleMapsApiKey: string }} */
  function loadConfig() {
    return {
      names: [...__configSnapshot.names],
      baseCps: [...__configSnapshot.baseCps],
      googleMapsApiKey: __configSnapshot.googleMapsApiKey,
    };
  }

  function schedulePersistConfig() {
    clearTimeout(persistConfigTimer);
    persistConfigTimer = setTimeout(() => void persistConfigToCloud(), 350);
  }

  async function persistConfigToCloud() {
    if (!supabaseClient) return;
    const {
      data: { user },
    } = await supabaseClient.auth.getUser();
    if (!user?.id) return;
    const { names, baseCps, googleMapsApiKey } = __configSnapshot;
    const { error } = await supabaseClient.from("agenda_user_settings").upsert(
      {
        user_id: user.id,
        payload: { names, baseCps, googleMapsApiKey },
        updated_at: new Date().toISOString(),
      },
      { onConflict: "user_id" }
    );
    if (error) console.error(error);
  }

  /** Lê config da nuvem; se vazia e existir cópia antiga em localStorage, envia e limpa o local. */
  async function hydrateConfigFromCloud() {
    if (!supabaseClient) return;
    const {
      data: { user },
    } = await supabaseClient.auth.getUser();
    if (!user?.id) return;
    const { data, error } = await supabaseClient
      .from("agenda_user_settings")
      .select("payload")
      .eq("user_id", user.id)
      .maybeSingle();
    if (error) throw error;
    const pl = data?.payload;
    if (pl && typeof pl === "object") {
      __configSnapshot = parseConfigFromRaw(pl);
      return;
    }
    const raw = localStorage.getItem(CONFIG_STORAGE_KEY);
    if (!raw) {
      __configSnapshot = defaultAgendaConfig();
      return;
    }
    try {
      __configSnapshot = parseConfigFromRaw(JSON.parse(raw));
    } catch {
      __configSnapshot = defaultAgendaConfig();
      return;
    }
    try {
      localStorage.removeItem(CONFIG_STORAGE_KEY);
    } catch {}
    await persistConfigToCloud();
  }

  /** @param {{ names: string[], baseCps?: string[], googleMapsApiKey?: string }} cfg */
  function saveConfig(cfg) {
    const names = cfg.names.slice(0, TECH_COUNT);
    while (names.length < TECH_COUNT) names.push("");
    const baseCps = Array.isArray(cfg.baseCps) ? cfg.baseCps.map(String).slice(0, TECH_COUNT) : [];
    while (baseCps.length < TECH_COUNT) baseCps.push("");
    const googleMapsApiKey = typeof cfg.googleMapsApiKey === "string" ? cfg.googleMapsApiKey.trim() : "";
    __configSnapshot = { names, baseCps, googleMapsApiKey };
    if (isLocalOnly()) {
      try {
        localStorage.setItem(CONFIG_STORAGE_KEY, JSON.stringify(__configSnapshot));
      } catch (e) {
        console.error(e);
      }
      return;
    }
    schedulePersistConfig();
  }

  /** @returns {{ index: number, name: string }[]} */
  function getVisibleTechnicians() {
    const { names } = loadConfig();
    const out = [];
    for (let i = 0; i < TECH_COUNT; i++) {
      const n = (names[i] || "").trim();
      if (n) out.push({ index: i, name: n });
    }
    return out;
  }

  function getTechDisplayName(techIndex) {
    const n = (loadConfig().names[techIndex] || "").trim();
    return n || `Posição ${techIndex + 1}`;
  }

  function normalizeNameKey(s) {
    return String(s || "")
      .trim()
      .toLowerCase()
      .normalize("NFD")
      .replace(/\p{M}/gu, "");
  }

  /** Nomes completos autorizados para aparelhos a gás (acentos ignorados). */
  const GAS_AUTHORIZED_NAME_KEYS = new Set(["joao rocha", "marcos correia"]);
  /** Ordem de desempate quando não há CP de referência para calcular distância. */
  const GAS_AUTH_FALLBACK_ORDER = ["joao rocha", "marcos correia"];

  function isGasAuthorizedTechName(name) {
    return GAS_AUTHORIZED_NAME_KEYS.has(normalizeNameKey(name));
  }

  function gasAuthorizedFallbackRank(name) {
    const k = normalizeNameKey(name);
    const i = GAS_AUTH_FALLBACK_ORDER.indexOf(k);
    return i >= 0 ? i : GAS_AUTH_FALLBACK_ORDER.length;
  }

  /** Próximo dia útil (seg–sáb) com slot extra livre para o técnico. */
  function findNextFreeExtraDay(techIndex, startIso) {
    let d = new Date(startIso + "T12:00:00");
    if (Number.isNaN(d.getTime())) return null;
    for (let k = 0; k < 120; k++) {
      const wd = d.getDay();
      if (wd === 0) {
        d.setDate(d.getDate() + 1);
        continue;
      }
      const iso = toISODate(d);
      if (!getBookingAt(iso, techIndex, 8)) {
        return { dateStr: iso, slotIndex: MORNING_EXTRA_ROW, period: "morning" };
      }
      if (!getBookingAt(iso, techIndex, 9)) {
        return { dateStr: iso, slotIndex: AFTERNOON_EXTRA_ROW, period: "afternoon" };
      }
      d.setDate(d.getDate() + 1);
    }
    return null;
  }

  /** CPs únicos com marcações desse técnico nesse dia. */
  function getUniquePostalCodesForTechOnDate(dateStr, techIndex) {
    const all = loadAll();
    const prefix = `${dateStr}|${techIndex}|`;
    const out = new Set();
    for (const key of Object.keys(all)) {
      if (!key.startsWith(prefix)) continue;
      const b = normalizeBooking(all[key]);
      if (b && isValidPostalPt(b.postalCode)) out.add(b.postalCode.trim());
    }
    return out;
  }

  /** Vagas livres nos slots fixos 9h–13h (armazenamento 0–3) e 14h–18h (4–7). */
  function countFreeFixedSlotsMorningAfternoon(dateStr, techIndex) {
    let morning = 0;
    let afternoon = 0;
    for (let s = 0; s <= 3; s++) {
      if (!getBookingAt(dateStr, techIndex, s)) morning++;
    }
    for (let s = 4; s <= 7; s++) {
      if (!getBookingAt(dateStr, techIndex, s)) afternoon++;
    }
    return { morning, afternoon };
  }

  function formatWizardSlotCountsHtml(dateStr, techIndex) {
    const { morning, afternoon } = countFreeFixedSlotsMorningAfternoon(dateStr, techIndex);
    return `9h–13h: <strong>${morning}</strong> vaga(s) livre(s) · 14h–18h: <strong>${afternoon}</strong> vaga(s) livre(s)`;
  }

  /** Primeiro slot fixo livre (0–7) ou null se cheio. */
  function findFirstFreeFixedStorageSlot(dateStr, techIndex) {
    for (let s = 0; s <= 7; s++) {
      if (!getBookingAt(dateStr, techIndex, s)) return s;
    }
    return null;
  }

  function findFirstFreeExtraOnDay(dateStr, techIndex) {
    if (!getBookingAt(dateStr, techIndex, 8)) return { row: MORNING_EXTRA_ROW, period: "morning" };
    if (!getBookingAt(dateStr, techIndex, 9)) return { row: AFTERNOON_EXTRA_ROW, period: "afternoon" };
    return null;
  }

  /** Próximo dia útil com vaga: prefere slots normais (9–18h); senão extra manhã/tarde. */
  function findFirstAvailableSlotPreferringFixed(techIndex, startIso) {
    let d = new Date(startIso + "T12:00:00");
    if (Number.isNaN(d.getTime())) return null;
    for (let k = 0; k < 120; k++) {
      if (d.getDay() === 0) {
        d.setDate(d.getDate() + 1);
        continue;
      }
      const iso = toISODate(d);
      const firstFixed = findFirstFreeFixedStorageSlot(iso, techIndex);
      if (firstFixed !== null) {
        return { dateStr: iso, slotIndex: storageToRow(firstFixed), isExtra: false };
      }
      if (!getBookingAt(iso, techIndex, 8)) {
        return { dateStr: iso, slotIndex: MORNING_EXTRA_ROW, isExtra: true, period: "morning" };
      }
      if (!getBookingAt(iso, techIndex, 9)) {
        return { dateStr: iso, slotIndex: AFTERNOON_EXTRA_ROW, isExtra: true, period: "afternoon" };
      }
      d.setDate(d.getDate() + 1);
    }
    return null;
  }

  /**
   * Wizard «por CP»: dia a dia, técnicos com marcação a ≤ ZONE_RADIUS_KM km;
   * por dia escolhe o técnico com menor distância à marcação mais próxima; prefere vaga em slot normal, senão extra (aviso separado).
   * @param {number[]} techIndices — índices permitidos (ex.: visíveis ou gás)
   */
  async function findBestZoneWizardMatch(customerCp, startIso, techIndices) {
    if (!(await geocodePostalPt(customerCp))) return null;

    const order = [...techIndices].sort((a, b) => a - b);
    let d = new Date(startIso + "T12:00:00");
    if (Number.isNaN(d.getTime())) return null;

    for (let k = 0; k < 120; k++) {
      if (d.getDay() === 0) {
        d.setDate(d.getDate() + 1);
        continue;
      }
      const dateStr = toISODate(d);
      /** @type {{ techIndex: number, bestKm: number, nearCp: string }[]} */
      const candidates = [];

      for (const techIndex of order) {
        const cps = getUniquePostalCodesForTechOnDate(dateStr, techIndex);
        if (cps.size === 0) continue;

        let bestKm = Infinity;
        let bestCp = "";
        for (const cp of cps) {
          const km = await postalDistanceKm(customerCp, cp);
          if (km === null) continue;
          if (km <= ZONE_RADIUS_KM && km < bestKm) {
            bestKm = km;
            bestCp = cp;
          }
        }
        if (!Number.isFinite(bestKm) || bestKm === Infinity) continue;
        candidates.push({ techIndex, bestKm, nearCp: bestCp });
      }

      candidates.sort((a, b) => a.bestKm - b.bestKm || a.techIndex - b.techIndex);

      for (const c of candidates) {
        const counts = countFreeFixedSlotsMorningAfternoon(dateStr, c.techIndex);
        const firstFixed = findFirstFreeFixedStorageSlot(dateStr, c.techIndex);
        if (firstFixed !== null) {
          return {
            kind: "zone_fixed",
            dateStr,
            techIndex: c.techIndex,
            slotIndex: storageToRow(firstFixed),
            km: c.bestKm,
            nearCp: c.nearCp,
            morningFree: counts.morning,
            afternoonFree: counts.afternoon,
          };
        }
        const extra = findFirstFreeExtraOnDay(dateStr, c.techIndex);
        if (extra) {
          return {
            kind: "zone_extra",
            dateStr,
            techIndex: c.techIndex,
            slotIndex: extra.row,
            km: c.bestKm,
            nearCp: c.nearCp,
            morningFree: counts.morning,
            afternoonFree: counts.afternoon,
            period: extra.period,
          };
        }
      }

      d.setDate(d.getDate() + 1);
    }
    return null;
  }

  /** CPs únicos por técnico extraídos das marcações já guardadas na agenda. */
  function collectPostalCodesByTechFromAgenda() {
    const all = loadAll();
    /** @type {Map<number, Set<string>>} */
    const map = new Map();
    for (const key of Object.keys(all)) {
      if (key === "__schema") continue;
      const parts = key.split("|");
      if (parts.length !== 3) continue;
      const techIndex = Number(parts[1]);
      if (!Number.isFinite(techIndex) || techIndex < 0 || techIndex >= TECH_COUNT) continue;
      const b = normalizeBooking(all[key]);
      if (!b || !isValidPostalPt(b.postalCode)) continue;
      const cp = b.postalCode.trim();
      if (!map.has(techIndex)) map.set(techIndex, new Set());
      map.get(techIndex).add(cp);
    }
    return map;
  }

  /**
   * Técnicos ordenados por distância do CP do cliente ao ponto mais próximo já presente na agenda
   * (CPs das marcações). Se um técnico ainda não tiver CP na agenda, usa-se o CP de referência em «Técnicos».
   */
  async function rankTechniciansByDistanceFromCustomerPostal(customerCp) {
    if (!(await geocodePostalPt(customerCp))) return [];

    const agendaCpsByTech = collectPostalCodesByTechFromAgenda();
    const { names, baseCps } = loadConfig();

    const out = [];
    for (let i = 0; i < TECH_COUNT; i++) {
      const n = (names[i] || "").trim();
      if (!n) continue;

      let bestKm = Infinity;
      const agendaSet = agendaCpsByTech.get(i);
      if (agendaSet && agendaSet.size > 0) {
        for (const cp of agendaSet) {
          const km = await postalDistanceKm(customerCp, cp);
          if (km !== null && km < bestKm) bestKm = km;
        }
      }
      if (!Number.isFinite(bestKm) || bestKm === Infinity) {
        const baseCp = (baseCps[i] || "").trim();
        if (isValidPostalPt(baseCp)) {
          const km = await postalDistanceKm(customerCp, baseCp);
          if (km !== null) bestKm = km;
        }
      }
      if (!Number.isFinite(bestKm) || bestKm === Infinity) continue;
      out.push({ index: i, name: n, km: bestKm });
    }
    out.sort((a, b) => a.km - b.km);
    return out;
  }

  function pad2(n) {
    return String(n).padStart(2, "0");
  }

  function isLunchRow(rowIndex) {
    return rowIndex === LUNCH_ROW_INDEX;
  }

  function isMorningExtraRow(rowIndex) {
    return rowIndex === MORNING_EXTRA_ROW;
  }

  function isAfternoonExtraRow(rowIndex) {
    return rowIndex === AFTERNOON_EXTRA_ROW;
  }

  /** Linha da grelha → armazenamento (0–7 fixos; 8 extra manhã; 9 extra tarde). */
  function rowToStorageSlot(rowIndex) {
    if (isLunchRow(rowIndex)) return -1;
    if (rowIndex === MORNING_EXTRA_ROW) return 8;
    if (rowIndex === AFTERNOON_EXTRA_ROW) return 9;
    if (rowIndex < MORNING_EXTRA_ROW) return rowIndex;
    if (rowIndex < AFTERNOON_EXTRA_ROW) return rowIndex - 2;
    return -1;
  }

  function storageToRow(storageSlot) {
    if (storageSlot < 0) return -1;
    if (storageSlot <= 3) return storageSlot;
    if (storageSlot <= 7) return storageSlot + 2;
    if (storageSlot === 8) return MORNING_EXTRA_ROW;
    if (storageSlot === 9) return AFTERNOON_EXTRA_ROW;
    return -1;
  }

  /** Marcação extra (sem horário; período manhã ou tarde). */
  function isExtraServiceSlot(slotIndex) {
    return slotIndex === MORNING_EXTRA_ROW || slotIndex === AFTERNOON_EXTRA_ROW;
  }

  /** Rótulo por linha da grelha. */
  function slotLabel(rowIndex) {
    if (rowIndex === MORNING_EXTRA_ROW) return "Extra — manhã";
    if (isLunchRow(rowIndex)) return "Almoço 13:00–14:00";
    if (rowIndex === AFTERNOON_EXTRA_ROW) return "Extra — tarde";
    if (rowIndex < MORNING_EXTRA_ROW) {
      const h = START_HOUR + rowIndex;
      return `${pad2(h)}:00–${pad2(h + 1)}:00`;
    }
    const h = 14 + (rowIndex - 6);
    return `${pad2(h)}:00–${pad2(h + 1)}:00`;
  }

  function formatManualWindow(booking) {
    if (!booking || !booking.timeFrom || !booking.timeTo) return "";
    return `${booking.timeFrom}–${booking.timeTo}`;
  }

  function parseTimeToMinutes(hhmm) {
    if (!hhmm || typeof hhmm !== "string") return NaN;
    const p = hhmm.split(":");
    const h = Number(p[0]);
    const m = Number(p[1] ?? 0);
    if (!Number.isFinite(h) || !Number.isFinite(m)) return NaN;
    return h * 60 + m;
  }

  /** Limites em minutos: armazenamento 0–3 manhã (9–13h); 4–7 tarde (14h–18h). */
  function storageSlotBoundsMinutes(storageIndex) {
    if (storageIndex < 0 || storageIndex >= BOOKABLE_STORAGE_SLOTS) return null;
    if (storageIndex <= 3) {
      const start = (START_HOUR + storageIndex) * 60;
      return { start, end: start + 60 };
    }
    const h = 14 + (storageIndex - 4);
    const start = h * 60;
    return { start, end: start + 60 };
  }

  /** Slots fixos (9h–13h e 14h–18h) cuja hora de 1h se sobrepõe a [timeFrom, timeTo). Índices = armazenamento 0–7. */
  function coveredFixedSlotIndices(timeFrom, timeTo) {
    const from = parseTimeToMinutes(timeFrom);
    const to = parseTimeToMinutes(timeTo);
    if (!Number.isFinite(from) || !Number.isFinite(to) || from >= to) return [];
    const out = [];
    for (let i = 0; i < BOOKABLE_STORAGE_SLOTS; i++) {
      const bounds = storageSlotBoundsMinutes(i);
      if (!bounds) continue;
      const { start, end } = bounds;
      if (from < end && to > start) out.push(i);
    }
    return out;
  }

  /**
   * Cobertura fixa ao começar em destStartStorage com a mesma duração (n slots de 1 h).
   * Não atravessa a pausa de almoço (último slot manhã 3 → primeiro tarde 4).
   */
  function computeFixedCoverageForRelocation(destStartStorage, slotCount) {
    if (slotCount < 1 || slotCount > 4) return null;
    const end = destStartStorage + slotCount - 1;
    if (destStartStorage < 0 || end > 7) return null;
    if (destStartStorage <= 3 && end > 3) return null;
    if (destStartStorage >= 4 && end > 7) return null;
    return Array.from({ length: slotCount }, (_, i) => destStartStorage + i);
  }

  function minutesToHHMM(totalMinutes) {
    const h = Math.floor(totalMinutes / 60) % 24;
    const m = Math.round(totalMinutes % 60);
    return `${pad2(h)}:${pad2(m)}`;
  }

  /** Início do 1.º slot e fim do último → intervalo guardado na marcação. */
  function timeRangeFromStorageIndices(indices) {
    const first = storageSlotBoundsMinutes(indices[0]);
    const last = storageSlotBoundsMinutes(indices[indices.length - 1]);
    if (!first || !last) return null;
    return {
      timeFrom: minutesToHHMM(first.start),
      timeTo: minutesToHHMM(last.end),
    };
  }

  function hhmmSlotStartForRow(rowIndex) {
    if (rowIndex < MORNING_EXTRA_ROW) return `${pad2(START_HOUR + rowIndex)}:00`;
    if (rowIndex > LUNCH_ROW_INDEX && rowIndex < AFTERNOON_EXTRA_ROW) return `${pad2(14 + (rowIndex - 6))}:00`;
    return "";
  }

  function hhmmSlotEndForRow(rowIndex) {
    if (rowIndex < MORNING_EXTRA_ROW) return `${pad2(START_HOUR + rowIndex + 1)}:00`;
    if (rowIndex > LUNCH_ROW_INDEX && rowIndex < AFTERNOON_EXTRA_ROW) return `${pad2(14 + (rowIndex - 6) + 1)}:00`;
    return "";
  }

  function newRangeId() {
    const c = globalThis.crypto;
    return c && typeof c.randomUUID === "function" ? c.randomUUID() : `r-${Date.now()}-${Math.random().toString(36).slice(2)}`;
  }

  function removeRangeFromStorage(dateStr, techIndex, rangeId) {
    if (!rangeId) return;
    const all = loadAll();
    let changed = false;
    const prefix = `${dateStr}|${techIndex}|`;
    for (const key of Object.keys(all)) {
      if (!key.startsWith(prefix)) continue;
      const b = all[key];
      if (b && b.rangeId === rangeId) {
        delete all[key];
        changed = true;
      }
    }
    if (changed) saveAll(all);
  }

  function setBookingsAcrossSlots(dateStr, techIndex, slotIndices, payload) {
    const all = loadAll();
    for (const idx of slotIndices) {
      all[keyFor(dateStr, techIndex, idx)] = { ...payload };
    }
    saveAll(all);
  }

  /**
   * @param {{ ignoreRangeId?: string, legacyEditSlot?: number }} [opts]
   */
  function wouldConflict(dateStr, techIndex, slotIndices, opts) {
    const ignoreRangeId = opts && opts.ignoreRangeId;
    const legacyEditSlot = opts && typeof opts.legacyEditSlot === "number" ? opts.legacyEditSlot : null;

    for (const idx of slotIndices) {
      const b = getBookingAt(dateStr, techIndex, idx);
      if (!b) continue;
      if (ignoreRangeId && b.rangeId === ignoreRangeId) continue;
      if (legacyEditSlot !== null && idx === legacyEditSlot && !b.rangeId) continue;
      return true;
    }
    return false;
  }

  const PT_POSTAL_RE = /^\d{4}-\d{3}$/;

  /**
   * Zonas de CP (4 primeiros dígitos) para cor na grelha e nos campos.
   * Prefixos repetidos em vários grupos: ganha o grupo de número mais baixo.
   */
  const POSTAL_PREFIX_GROUPS_RAW = [
    ["2705", "2710", "2715"],
    ["2735", "2605", "2745", "2635", "2710", "2725"],
    ["2645", "2750", "2755", "2765", "2775", "2780", "2785"],
    ["1495", "2730", "2740", "2760", "2790"],
    ["2610", "2700"],
    ["1675", "1686", "2620", "2621", "2675", "2605"],
    ["2600", "2615", "2626", "2690", "2695"],
    ["1800", "1885", "1900", "1950", "2680", "2685", "1600"],
    ["2670", "2660", "2671"],
    ["1500", "1070", "1050"],
    ["1200", "1250", "1300", "1350", "1400"],
    ["1000", "1100", "1170"],
  ];

  const postalPrefixToGroup = (() => {
    const m = new Map();
    for (let i = 0; i < POSTAL_PREFIX_GROUPS_RAW.length; i++) {
      const gi = i + 1;
      for (const x of POSTAL_PREFIX_GROUPS_RAW[i]) {
        const p = String(x).replace(/\D/g, "").slice(0, 4);
        if (p.length === 4 && !m.has(p)) m.set(p, gi);
      }
    }
    return m;
  })();

  function isValidPostalPt(value) {
    return PT_POSTAL_RE.test(String(value || "").trim());
  }

  /** @returns {number | null} índice do grupo 1–12 */
  function postalGroupIndexFromCp(cp) {
    const digits = String(cp || "").replace(/\D/g, "");
    if (digits.length < 4) return null;
    return postalPrefixToGroup.get(digits.slice(0, 4)) ?? null;
  }

  /** Classes CSS para inputs e etiquetas (legenda). */
  function postalZoneClassesForCp(cp) {
    const raw = String(cp || "").trim();
    const digits = raw.replace(/\D/g, "");
    if (digits.length < 4) return ["postal-zone--none"];
    const prefix = digits.slice(0, 4);
    const gi = postalPrefixToGroup.get(prefix);
    if (gi) return [`postal-zone--g${gi}`];
    if (digits.length === 7 && PT_POSTAL_RE.test(raw)) return ["postal-zone--unmapped"];
    return ["postal-zone--none"];
  }

  /** Classes no botão da célula (marcação ocupada) — cor em toda a célula. */
  function slotZoneClassesForCp(cp) {
    const raw = String(cp || "").trim();
    const digits = raw.replace(/\D/g, "");
    if (digits.length < 4) return [];
    const prefix = digits.slice(0, 4);
    const gi = postalPrefixToGroup.get(prefix);
    if (gi) return [`slot--zone-g${gi}`];
    if (digits.length === 7 && PT_POSTAL_RE.test(raw)) return ["slot--zone-unmapped"];
    return [];
  }

  function stripSlotZoneClasses(btn) {
    [...btn.classList].forEach((c) => {
      if (c.startsWith("slot--zone-")) btn.classList.remove(c);
    });
  }

  function applySlotZoneClasses(btn, postalCode) {
    stripSlotZoneClasses(btn);
    for (const c of slotZoneClassesForCp(postalCode)) btn.classList.add(c);
  }

  function applyPostalZoneClassesToInput(el) {
    if (!el) return;
    for (let i = 1; i <= 12; i++) el.classList.remove(`postal-zone--g${i}`);
    el.classList.remove("postal-zone--none", "postal-zone--unmapped");
    for (const c of postalZoneClassesForCp(el.value)) el.classList.add(c);
  }

  /** Normaliza marcações antigas (client / notes) e campos em falta. */
  function normalizeBooking(b) {
    if (!b) return null;
    let extraPeriod = b.extraPeriod;
    if (extraPeriod !== "morning" && extraPeriod !== "afternoon") {
      extraPeriod = undefined;
    }
    const timeFrom = b.timeFrom;
    const timeTo = b.timeTo;
    if (!extraPeriod && (timeFrom || timeTo)) {
      const from = parseTimeToMinutes(timeFrom);
      if (Number.isFinite(from)) extraPeriod = from < 13 * 60 ? "morning" : "afternoon";
    }
    return {
      processNo: String(b.processNo ?? b.processNumber ?? "").trim(),
      clientName: String(b.clientName ?? b.client ?? "").trim(),
      phone: String(b.phone ?? "").trim(),
      postalCode: String(b.postalCode ?? "").trim(),
      device: String(b.device ?? "").trim(),
      observations: String(b.observations ?? b.notes ?? "").trim(),
      timeFrom,
      timeTo,
      extraPeriod,
      rangeId: b.rangeId,
      createdAt: b.createdAt ?? Date.now(),
      gasAppliance: Boolean(b.gasAppliance),
    };
  }

  function buildBookingTooltip(b, slotIndex) {
    const n = normalizeBooking(b);
    if (!n) return "";
    const lines = [
      `Processo: ${n.processNo || "—"}`,
      `Cliente: ${n.clientName || "—"}`,
      `Telefone: ${n.phone || "—"}`,
      `Código postal: ${n.postalCode || "—"}`,
      `Aparelho: ${n.device || "—"}`,
      `Observações: ${n.observations || "—"}`,
    ];
    if (n.gasAppliance) {
      lines.push("Aparelho a gás: sim");
    }
    if (n.extraPeriod === "morning" || n.extraPeriod === "afternoon") {
      lines.push(`Período: ${n.extraPeriod === "morning" ? "Manhã" : "Tarde"}`);
    } else if (formatManualWindow(n)) {
      lines.push(`Horário: ${formatManualWindow(n)}`);
      const cov = coveredFixedSlotIndices(n.timeFrom, n.timeTo);
      if (cov.length > 1) lines.push(`Duração: ${cov.length} slots (${cov.length} h)`);
    }
    return lines.join("\n");
  }

  /**
   * @param {object} booking
   * @param {number} rowIndex — linha da grelha (não armazenamento)
   * @param {boolean} isExtraCell
   * @returns {{ kind: "single", span: number } | { kind: "span", span: number, segment: "first" | "middle" | "last" } | null}
   */
  function getFixedMultiSlotSegment(booking, rowIndex, isExtraCell) {
    if (isExtraCell || !booking) return null;
    const n = normalizeBooking(booking);
    if (!n.timeFrom || !n.timeTo) return null;
    const cov = coveredFixedSlotIndices(n.timeFrom, n.timeTo);
    if (cov.length === 0) return null;
    if (cov.length === 1) return { kind: "single", span: 1 };
    const st = rowToStorageSlot(rowIndex);
    if (!Number.isFinite(st) || st < 0) return null;
    if (!cov.includes(st)) return null;
    const startSt = Math.min(...cov);
    const endSt = Math.max(...cov);
    const span = cov.length;
    if (st === startSt) return { kind: "span", span, segment: "first" };
    if (st === endSt) return { kind: "span", span, segment: "last" };
    return { kind: "span", span, segment: "middle" };
  }

  function spanLabelPt(span) {
    if (span === 2) return "Duplo";
    if (span === 3) return "Triplo";
    if (span === 4) return "Quádruplo";
    return "";
  }

  /**
   * @param {{ kind: "span", span: number, segment: "middle" | "last" }} seg
   */
  function fillBookingCellContinuation(btn, booking, slotIndex, seg) {
    btn.classList.add("slot--busy-span-continuation");
    const mark = document.createElement("span");
    mark.className = "slot__span-continue-mark";
    mark.setAttribute("aria-hidden", "true");
    mark.innerHTML = "&nbsp;";
    btn.appendChild(mark);
    const sr = document.createElement("span");
    sr.className = "visually-hidden";
    sr.textContent = `Mesmo serviço (${spanLabelPt(seg.span)}). Ver primeira linha do bloco.`;
    btn.appendChild(sr);
    btn.title = buildBookingTooltip(booking, slotIndex);
  }

  /**
   * @param {{ kind: "single", span: number } | { kind: "span", span: number, segment: "first" | "middle" | "last" } | null} [multiSeg]
   */
  function fillBookingCell(btn, booking, slotIndex, multiSeg) {
    const n = normalizeBooking(booking);
    btn.textContent = "";
    const seg = multiSeg || { kind: "single", span: 1 };

    if (seg.kind === "span" && seg.segment !== "first") {
      fillBookingCellContinuation(btn, booking, slotIndex, seg);
      return;
    }

    if (isExtraServiceSlot(slotIndex)) {
      const tag = document.createElement("span");
      tag.className = "slot__tag slot__tag--extra";
      tag.textContent = "EXTRA";
      btn.appendChild(tag);
    } else if (seg.kind === "span" && seg.span > 1) {
      const tag = document.createElement("span");
      tag.className = "slot__tag slot__tag--multi";
      tag.textContent = spanLabelPt(seg.span);
      btn.appendChild(tag);
    }

    const l1 = document.createElement("span");
    l1.className = "slot__line slot__line--primary";
    l1.textContent = n.processNo || "—";
    const l2 = document.createElement("span");
    l2.className = "slot__line slot__line--secondary";
    const cpRaw = n.postalCode || "";
    const cpDisplay = cpRaw || "—";
    const dev = n.device || "—";
    const cpSpan = document.createElement("span");
    cpSpan.className = "slot__cp";
    cpSpan.textContent = cpDisplay;
    const sep = document.createTextNode(" · ");
    const devSpan = document.createElement("span");
    devSpan.className = "slot__line-dev";
    devSpan.textContent = dev;
    l2.appendChild(cpSpan);
    l2.appendChild(sep);
    l2.appendChild(devSpan);
    btn.appendChild(l1);
    btn.appendChild(l2);
    btn.title = buildBookingTooltip(booking, slotIndex);
  }

  function atNoon(d) {
    const x = new Date(d);
    x.setHours(12, 0, 0, 0);
    return x;
  }

  function toISODate(d) {
    const x = atNoon(d);
    const y = x.getFullYear();
    const m = String(x.getMonth() + 1).padStart(2, "0");
    const day = String(x.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
  }

  function todayISODate() {
    return toISODate(new Date());
  }

  function mondayOfWeekContaining(d) {
    const x = atNoon(d);
    const day = x.getDay();
    const diff = day === 0 ? -6 : 1 - day;
    x.setDate(x.getDate() + diff);
    return x;
  }

  /**
   * @param {number} year
   * @param {number} monthIndex 0–11
   * @returns {{ monday: Date, days: Date[] }[]}
   */
  function weeksOverlappingMonth(year, monthIndex) {
    const first = atNoon(new Date(year, monthIndex, 1));
    const last = atNoon(new Date(year, monthIndex + 1, 0));
    let mon = mondayOfWeekContaining(first);
    const endMon = mondayOfWeekContaining(last);
    const weeks = [];

    while (mon.getTime() <= endMon.getTime()) {
      const days = [];
      for (let i = 0; i < 6; i++) {
        const day = new Date(mon);
        day.setDate(mon.getDate() + i);
        days.push(atNoon(day));
      }
      weeks.push({ monday: new Date(mon), days });
      const next = new Date(mon);
      next.setDate(mon.getDate() + 7);
      mon = atNoon(next);
    }
    return weeks;
  }

  function formatShortRange(days) {
    const a = days[0];
    const b = days[5];
    const y = a.getFullYear();
    if (a.getMonth() === b.getMonth() && a.getFullYear() === b.getFullYear()) {
      return `${a.getDate()}–${b.getDate()} ${a.toLocaleDateString("pt-PT", { month: "long" })} ${y}`;
    }
    const left = a.toLocaleDateString("pt-PT", { day: "numeric", month: "short" });
    const right = b.toLocaleDateString("pt-PT", { day: "numeric", month: "short", year: "numeric" });
    return `${left} – ${right}`;
  }

  function weekContainsTodayIso(week) {
    const t = todayISODate();
    return week.days.some((d) => toISODate(d) === t);
  }

  function weekContainsIsoDate(week, iso) {
    return week.days.some((d) => toISODate(d) === iso);
  }

  /** Migra chaves antigas (8 slots 9–17h contínuos) para o modelo com pausa 13–14h (7 slots + extras). Marcações só no antigo slot 13–14 são omitidas. */
  function migrateSlotKeysForLunch(parsed) {
    const out = { __schema: 2 };
    for (const [k, v] of Object.entries(parsed)) {
      if (k === "__schema") continue;
      const parts = k.split("|");
      if (parts.length !== 3) continue;
      const [d, t, s] = parts;
      const si = Number(s);
      if (!Number.isFinite(si)) continue;
      if (si <= 3) {
        out[k] = v;
        continue;
      }
      if (si === 4) continue;
      if (si >= 5 && si <= 7) {
        out[`${d}|${t}|${si - 1}`] = v;
        continue;
      }
      if (si >= 8) out[k] = v;
    }
    return out;
  }

  function loadAll() {
    if (!__bookingStore || typeof __bookingStore !== "object") {
      __bookingStore = { __schema: 2 };
    }
    if (__bookingStore.__schema !== 2) {
      __bookingStore = migrateSlotKeysForLunch(__bookingStore);
      __bookingStore.__schema = 2;
    }
    return __bookingStore;
  }

  function saveAll(data) {
    if (!data || typeof data !== "object") return;
    __bookingStore = data;
    if (__bookingStore.__schema !== 2) __bookingStore.__schema = 2;
    schedulePersistBookings();
  }

  /**
   * Modo só-browser: só em localhost, sem credenciais válidas, e localOnly.
   * Com supabaseUrl + chave → sempre nuvem. Em produção (ex. Netlify) nunca modo só-browser.
   */
  function isLocalOnly() {
    const cfg = window.AGENDA_CONFIG;
    if (!cfg) return false;
    const key = cfg.supabaseKey ? String(cfg.supabaseKey).trim() : "";
    const hasCloud = Boolean(cfg.supabaseUrl && key && !key.includes("COLOCA_AQUI"));
    if (hasCloud) return false;
    try {
      const host = globalThis.location?.hostname || "";
      const allowLocalMode =
        host === "localhost" || host === "127.0.0.1" || host === "[::1]";
      if (!allowLocalMode) return false;
    } catch {
      /* ignore */
    }
    return cfg.localOnly === true;
  }

  function schedulePersistBookings() {
    clearTimeout(persistBookingsTimer);
    persistBookingsTimer = setTimeout(() => {
      persistBookingsTimer = null;
      void persistBookings();
    }, 450);
  }

  function persistBookingsToLocal() {
    try {
      const all = { ...__bookingStore };
      localStorage.setItem(STORAGE_KEY, JSON.stringify(all));
    } catch (e) {
      console.error(e);
    }
  }

  function loadBookingsFromLocal() {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      __bookingStore = { __schema: 2 };
      return;
    }
    try {
      let parsed = JSON.parse(raw);
      if (typeof parsed !== "object" || parsed === null) {
        __bookingStore = { __schema: 2 };
        return;
      }
      if (parsed.__schema !== 2) {
        parsed = migrateSlotKeysForLunch(parsed);
      }
      __bookingStore = parsed;
      __bookingStore.__schema = 2;
    } catch {
      __bookingStore = { __schema: 2 };
    }
  }

  async function persistBookings() {
    if (isLocalOnly()) {
      persistBookingsToLocal();
      return;
    }
    persistBookingsInFlight++;
    try {
      await persistBookingsToCloud();
    } finally {
      persistBookingsInFlight--;
    }
  }

  async function persistBookingsToCloud() {
    if (!supabaseClient) return;
    const all = { ...__bookingStore };
    delete all.__schema;
    const keys = Object.keys(all).filter((k) => k.split("|").length === 3);
    const keySet = new Set(keys);

    const { data: existing, error: errSel } = await supabaseClient
      .from("agenda_bookings")
      .select("date_str, tech_index, storage_slot");
    if (errSel) {
      console.error(errSel);
      return;
    }
    for (const r of existing || []) {
      const k = `${r.date_str}|${r.tech_index}|${r.storage_slot}`;
      if (!keySet.has(k)) {
        const { error: errDel } = await supabaseClient
          .from("agenda_bookings")
          .delete()
          .eq("date_str", r.date_str)
          .eq("tech_index", r.tech_index)
          .eq("storage_slot", r.storage_slot);
        if (errDel) console.error(errDel);
      }
    }

    const rows = keys.map((k) => {
      const [d, t, s] = k.split("|");
      return {
        date_str: d,
        tech_index: Number(t),
        storage_slot: Number(s),
        payload: all[k],
        updated_at: new Date().toISOString(),
      };
    });
    if (rows.length === 0) return;
    const { error: errUp } = await supabaseClient.from("agenda_bookings").upsert(rows, {
      onConflict: "date_str,tech_index,storage_slot",
    });
    if (errUp) console.error(errUp);
  }

  async function hydrateBookingsFromCloud() {
    if (!supabaseClient) return;
    const { data, error } = await supabaseClient
      .from("agenda_bookings")
      .select("date_str, tech_index, storage_slot, payload");
    if (error) throw error;
    const obj = { __schema: 2 };
    for (const row of data || []) {
      const key = `${row.date_str}|${row.tech_index}|${row.storage_slot}`;
      obj[key] = row.payload;
    }
    __bookingStore = obj;
  }

  function scheduleRealtimeGridRefresh() {
    clearTimeout(realtimeRenderTimer);
    realtimeRenderTimer = setTimeout(() => {
      realtimeRenderTimer = null;
      renderGrid({});
    }, 100);
  }

  function bookingKeyFromRow(row) {
    if (!row || row.date_str == null || row.tech_index == null || row.storage_slot == null) return null;
    const t = Number(row.tech_index);
    const s = Number(row.storage_slot);
    if (!Number.isFinite(t) || !Number.isFinite(s)) return null;
    return `${row.date_str}|${t}|${s}`;
  }

  /** @param {{ eventType: string, new: Record<string, unknown> | null, old: Record<string, unknown> | null }} payload */
  function applyRealtimeBookingChange(payload) {
    if (!__bookingStore || typeof __bookingStore !== "object") return;
    __bookingStore.__schema = 2;
    const ev = payload.eventType;
    if (ev === "DELETE") {
      const row = payload.old;
      const key = row ? bookingKeyFromRow(row) : null;
      if (key) delete __bookingStore[key];
      scheduleRealtimeGridRefresh();
      return;
    }
    const row = payload.new;
    const key = row ? bookingKeyFromRow(row) : null;
    if (!key) return;
    let pl = row.payload;
    if (typeof pl === "string") {
      try {
        pl = JSON.parse(pl);
      } catch {
        return;
      }
    }
    __bookingStore[key] = pl;
    scheduleRealtimeGridRefresh();
  }

  /** @param {{ eventType: string, new: Record<string, unknown> | null }} payload */
  function applyRealtimeUserSettingsChange(payload) {
    if (payload.eventType === "DELETE") return;
    const row = payload.new;
    if (row && row.payload != null) {
      __configSnapshot = parseConfigFromRaw(row.payload);
      scheduleRealtimeGridRefresh();
    }
  }

  function unsubscribeRealtimeAgenda() {
    clearTimeout(realtimeRenderTimer);
    realtimeRenderTimer = null;
    if (agendaPollIntervalId != null) {
      clearInterval(agendaPollIntervalId);
      agendaPollIntervalId = null;
    }
    try {
      if (supabaseClient && realtimeAgendaChannels.length) {
        for (const ch of realtimeAgendaChannels) {
          void supabaseClient.removeChannel(ch);
        }
      }
    } catch (e) {
      console.error(e);
    }
    realtimeAgendaChannels = [];
  }

  /** Postgres Changes + RLS: o servidor Realtime precisa do JWT da sessão. */
  function setRealtimeAuthFromSession(session) {
    if (!supabaseClient || !session?.access_token) return;
    const rt = supabaseClient.realtime;
    if (rt && typeof rt.setAuth === "function") {
      try {
        rt.setAuth(session.access_token);
      } catch (e) {
        console.error(e);
      }
    }
  }

  /** Releitura da nuvem (fallback leve quando o WebSocket falha). */
  async function pullBookingsFromCloudAndRender() {
    if (!supabaseClient || isLocalOnly() || document.hidden) return;
    if (persistBookingsTimer != null || persistBookingsInFlight > 0) return;
    try {
      await hydrateBookingsFromCloud();
      renderGrid({});
    } catch (e) {
      console.error(e);
    }
  }

  async function subscribeRealtimeAgenda() {
    if (!supabaseClient || isLocalOnly()) return;
    unsubscribeRealtimeAgenda();

    const {
      data: { session },
    } = await supabaseClient.auth.getSession();
    if (!session?.access_token) return;
    setRealtimeAuthFromSession(session);

    const {
      data: { user },
    } = await supabaseClient.auth.getUser();
    if (!user?.id) return;

    const uid = user.id;
    const filterUser = `user_id=eq.${uid}`;
    const suffix = `${uid}:${Date.now()}`;

    /** Dois canais evitam conflitos entre filtro / tabelas em alguns proxies. */
    const chBookings = supabaseClient
      .channel(`agenda-rt-bookings:${suffix}`)
      .on(
        "postgres_changes",
        { event: "*", schema: "public", table: "agenda_bookings" },
        (payload) => applyRealtimeBookingChange(payload)
      )
      .subscribe((status, err) => {
        if (status === "CHANNEL_ERROR" || status === "TIMED_OUT") {
          console.error("Realtime agenda (marcações):", err || status);
        }
      });

    const chSettings = supabaseClient
      .channel(`agenda-rt-settings:${suffix}`)
      .on(
        "postgres_changes",
        {
          event: "*",
          schema: "public",
          table: "agenda_user_settings",
          filter: filterUser,
        },
        (payload) => applyRealtimeUserSettingsChange(payload)
      )
      .subscribe((status, err) => {
        if (status === "CHANNEL_ERROR" || status === "TIMED_OUT") {
          console.error("Realtime agenda (definições):", err || status);
        }
      });

    realtimeAgendaChannels = [chBookings, chSettings];

    agendaPollIntervalId = window.setInterval(() => {
      void pullBookingsFromCloudAndRender();
    }, 28000);
  }

  /** Se a nuvem está vazia mas ainda há dados antigos no browser, envia uma vez e limpa o local. */
  async function maybeMigrateLocalStorageToCloud() {
    if (!supabaseClient) return;
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    let parsed;
    try {
      parsed = JSON.parse(raw);
    } catch {
      return;
    }
    if (typeof parsed !== "object" || parsed === null) return;
    if (parsed.__schema !== 2) {
      parsed = migrateSlotKeysForLunch(parsed);
    }
    const { count, error: cErr } = await supabaseClient
      .from("agenda_bookings")
      .select("*", { count: "exact", head: true });
    if (cErr) return;
    if ((count ?? 0) > 0) return;

    const copy = { ...parsed };
    delete copy.__schema;
    __bookingStore = { __schema: 2 };
    Object.assign(__bookingStore, copy);
    __bookingStore.__schema = 2;
    try {
      localStorage.removeItem(STORAGE_KEY);
    } catch {}
    await persistBookingsToCloud();
  }

  function keyFor(dateStr, techIndex, storageSlot) {
    return `${dateStr}|${techIndex}|${storageSlot}`;
  }

  /** @param {number} storageSlot — 0–6 (fixos) ou 8–9 (extras) */
  function getBookingAt(dateStr, techIndex, storageSlot) {
    const all = loadAll();
    return all[keyFor(dateStr, techIndex, storageSlot)] || null;
  }

  /** @param {number} storageSlot */
  function setBookingAt(dateStr, techIndex, storageSlot, booking) {
    const all = loadAll();
    const k = keyFor(dateStr, techIndex, storageSlot);
    if (booking) all[k] = booking;
    else delete all[k];
    saveAll(all);
  }

  /** Limites [start,end) em minutos para linhas de slots fixos (não extra nem almoço). */
  function rowToTimeBoundsMinutes(rowIndex) {
    if (isLunchRow(rowIndex) || isMorningExtraRow(rowIndex) || isAfternoonExtraRow(rowIndex)) return null;
    if (rowIndex < MORNING_EXTRA_ROW) {
      const h = START_HOUR + rowIndex;
      return { start: h * 60, end: (h + 1) * 60 };
    }
    if (rowIndex > LUNCH_ROW_INDEX && rowIndex < AFTERNOON_EXTRA_ROW) {
      const h = 14 + (rowIndex - 6);
      return { start: h * 60, end: (h + 1) * 60 };
    }
    return null;
  }

  /**
   * @returns {{ booking: object | null, extraStorage: number | null, isExtra: boolean }}
   */
  function getCellDisplayBooking(dateStr, techIndex, rowIndex, inMonth) {
    if (!inMonth || isLunchRow(rowIndex)) return { booking: null, extraStorage: null, isExtra: false };
    if (rowIndex === MORNING_EXTRA_ROW) {
      const b = getBookingAt(dateStr, techIndex, 8);
      return b
        ? { booking: b, extraStorage: 8, isExtra: true }
        : { booking: null, extraStorage: null, isExtra: false };
    }
    if (rowIndex === AFTERNOON_EXTRA_ROW) {
      const b = getBookingAt(dateStr, techIndex, 9);
      return b
        ? { booking: b, extraStorage: 9, isExtra: true }
        : { booking: null, extraStorage: null, isExtra: false };
    }
    const st = rowToStorageSlot(rowIndex);
    const fixed = getBookingAt(dateStr, techIndex, st);
    return { booking: fixed, extraStorage: null, isExtra: false };
  }

  function isCellFreeForRelocation(iso, tech, s) {
    if (s === MORNING_EXTRA_ROW) return !getBookingAt(iso, tech, 8);
    if (s === AFTERNOON_EXTRA_ROW) return !getBookingAt(iso, tech, 9);
    return !getBookingAt(iso, tech, rowToStorageSlot(s));
  }

  const monthInput = document.getElementById("monthInput");
  const btnPrevMonth = document.getElementById("btnPrevMonth");
  const btnNextMonth = document.getElementById("btnNextMonth");
  const btnToday = document.getElementById("btnToday");
  const btnConfigTech = document.getElementById("btnConfigTech");
  const weekTabs = document.getElementById("weekTabs");
  const scheduleHead = document.getElementById("scheduleHead");
  const gridBody = document.getElementById("gridBody");
  const tableWrap = document.getElementById("tableWrap");
  const emptyTechState = document.getElementById("emptyTechState");
  const configModal = document.getElementById("configModal");
  const configForm = document.getElementById("configForm");
  const techNamesFields = document.getElementById("techNamesFields");
  const btnConfigClose = document.getElementById("btnConfigClose");

  const modal = document.getElementById("modal");
  const modalForm = document.getElementById("modalForm");
  const modalTitle = document.getElementById("modalTitle");
  const modalContext = document.getElementById("modalContext");
  const timeRangeFields = document.getElementById("timeRangeFields");
  const timeRangeHint = document.getElementById("timeRangeHint");
  const fieldTimeFrom = document.getElementById("fieldTimeFrom");
  const fieldTimeTo = document.getElementById("fieldTimeTo");
  const fieldProcess = document.getElementById("fieldProcess");
  const fieldClientName = document.getElementById("fieldClientName");
  const fieldPhone = document.getElementById("fieldPhone");
  const fieldPostal = document.getElementById("fieldPostal");
  const fieldDevice = document.getElementById("fieldDevice");
  const fieldObservations = document.getElementById("fieldObservations");
  const btnDelete = document.getElementById("btnDelete");
  const btnRelocateBooking = document.getElementById("btnRelocateBooking");
  const btnClose = document.getElementById("btnClose");
  const extraMetaFields = document.getElementById("extraMetaFields");
  const fieldExtraDate = document.getElementById("fieldExtraDate");
  const fieldExtraTech = document.getElementById("fieldExtraTech");
  const fieldExtraPeriodMorning = document.getElementById("fieldExtraPeriodMorning");
  const fieldExtraPeriodAfternoon = document.getElementById("fieldExtraPeriodAfternoon");
  const btnScheduleByPostal = document.getElementById("btnScheduleByPostal");
  const dialogScheduleByPostal = document.getElementById("dialogScheduleByPostal");
  const fieldWizardPostal = document.getElementById("fieldWizardPostal");
  const fieldWizardGas = document.getElementById("fieldWizardGas");
  const btnWizardSearch = document.getElementById("btnWizardSearch");
  const btnWizardContinue = document.getElementById("btnWizardContinue");
  const btnWizardClose = document.getElementById("btnWizardClose");
  const wizardTechResults = document.getElementById("wizardTechResults");
  const fieldSlotSpanGroup = document.getElementById("fieldSlotSpanGroup");

  /** Login e Supabase primeiro — se algo abaixo no script falhar, o botão «Entrar» continua a funcionar. */
  void startAgenda();

  /** @type {{ dateStr: string, techIndex: number, slotIndex: number, isGas: boolean } | null} */
  let wizardSuggest = null;

  function getExtraPeriodFromRadio() {
    if (fieldExtraPeriodAfternoon && fieldExtraPeriodAfternoon.checked) return "afternoon";
    return "morning";
  }

  function setExtraPeriodRadio(period) {
    if (period === "afternoon") {
      if (fieldExtraPeriodAfternoon) fieldExtraPeriodAfternoon.checked = true;
    } else if (fieldExtraPeriodMorning) {
      fieldExtraPeriodMorning.checked = true;
    }
  }

  function syncExtraSlotFromPeriodForButton() {
    if (!editContext || !editContext.extraFromButton) return;
    editContext.slotIndex = getExtraPeriodFromRadio() === "afternoon" ? AFTERNOON_EXTRA_ROW : MORNING_EXTRA_ROW;
  }

  function setTimeFieldsRequired(on) {
    if (fieldTimeFrom) fieldTimeFrom.required = Boolean(on);
    if (fieldTimeTo) fieldTimeTo.required = Boolean(on);
  }

  /**
   * @type {{
   *   dateStr: string,
   *   techIndex: number,
   *   slotIndex: number,
   *   existing: object | null,
   *   extraFromButton?: boolean,
   *   skipRemoveExisting?: boolean
   * } | null }
   */
  let editContext = null;

  /**
   * Snapshot ao escolher «Alterar» no aviso de zona: permite escolher célula na grelha antes de guardar.
   * @type {{
   *   dateStr: string,
   *   techIndex: number,
   *   slotIndex: number,
   *   existing: object | null,
   *   form: { processNo: string, clientName: string, phone: string, postalCode: string, device: string, observations: string, timeFrom: string, timeTo: string },
   *   extraFromButton?: boolean
   * } | null}
   */
  let relocationStaging = null;
  let relocationPickMode = false;

  /** Arrastar marcação na grelha (sem modal aberto). */
  let dragStaging = null;

  function maxValidSpanForAnchor(anchorStorage) {
    if (anchorStorage < 0 || anchorStorage > 7) return 1;
    for (let s = 4; s >= 1; s--) {
      if (computeFixedCoverageForRelocation(anchorStorage, s)) return s;
    }
    return 1;
  }

  function getSlotSpanFromRadios() {
    const el = document.querySelector('input[name="slotSpan"]:checked');
    const n = el ? Number(el.value) : 1;
    return n >= 1 && n <= 4 ? n : 1;
  }

  function setSlotSpanRadios(span) {
    const s = Math.min(4, Math.max(1, Math.round(span)));
    const r = document.querySelector(`input[name="slotSpan"][value="${s}"]`);
    if (r) r.checked = true;
  }

  function getCoverageAnchorStorageForModal() {
    if (!editContext || isExtraServiceSlot(editContext.slotIndex)) return 0;
    const tf = fieldTimeFrom && fieldTimeFrom.value;
    const tt = fieldTimeTo && fieldTimeTo.value;
    if (tf && tt && tf < tt) {
      const cov = coveredFixedSlotIndices(tf, tt);
      if (cov.length > 0) return Math.min(...cov);
    }
    return rowToStorageSlot(editContext.slotIndex);
  }

  function syncTimesFromAnchorAndSpan() {
    if (!editContext || isExtraServiceSlot(editContext.slotIndex)) return;
    const anchor = getCoverageAnchorStorageForModal();
    const span = getSlotSpanFromRadios();
    const cov = computeFixedCoverageForRelocation(anchor, span);
    if (!cov) {
      const maxS = maxValidSpanForAnchor(anchor);
      window.alert(
        `Esta duração não cabe a partir do slot inicial (máximo ${maxS} h sem atravessar a pausa 13h–14h).`
      );
      setSlotSpanRadios(maxS);
      const cov2 = computeFixedCoverageForRelocation(anchor, maxS);
      if (!cov2) return;
      const tr = timeRangeFromStorageIndices(cov2);
      if (tr && fieldTimeFrom && fieldTimeTo) {
        fieldTimeFrom.value = tr.timeFrom;
        fieldTimeTo.value = tr.timeTo;
      }
      return;
    }
    const tr = timeRangeFromStorageIndices(cov);
    if (tr && fieldTimeFrom && fieldTimeTo) {
      fieldTimeFrom.value = tr.timeFrom;
      fieldTimeTo.value = tr.timeTo;
    }
  }

  function syncSlotSpanRadiosFromTimeFields() {
    if (!editContext || isExtraServiceSlot(editContext.slotIndex)) return;
    const tf = fieldTimeFrom && fieldTimeFrom.value;
    const tt = fieldTimeTo && fieldTimeTo.value;
    if (!tf || !tt || tf >= tt) return;
    const cov = coveredFixedSlotIndices(tf, tt);
    if (cov.length < 1 || cov.length > 4) return;
    setSlotSpanRadios(cov.length);
  }

  function wireSlotSpanControls() {
    document.querySelectorAll('input[name="slotSpan"]').forEach((el) => {
      el.addEventListener("change", () => syncTimesFromAnchorAndSpan());
    });
    if (fieldTimeFrom) fieldTimeFrom.addEventListener("input", syncSlotSpanRadiosFromTimeFields);
    if (fieldTimeTo) fieldTimeTo.addEventListener("input", syncSlotSpanRadiosFromTimeFields);
  }

  /**
   * Snapshot para recolha / arrastar, a partir da marcação na célula.
   * @param {string} dateStr
   * @param {number} techIndex
   * @param {number} slotIndex
   * @param {object} booking
   */
  function buildStagingFromGridBooking(dateStr, techIndex, slotIndex, booking) {
    const b = normalizeBooking(booking);
    return {
      dateStr,
      techIndex,
      slotIndex,
      existing: booking,
      extraFromButton: false,
      form: {
        processNo: b.processNo,
        clientName: b.clientName,
        phone: b.phone,
        postalCode: b.postalCode,
        device: b.device,
        observations: b.observations,
        timeFrom: b.timeFrom || hhmmSlotStartForRow(slotIndex),
        timeTo: b.timeTo || hhmmSlotEndForRow(slotIndex),
        extraPeriod: b.extraPeriod === "afternoon" ? "afternoon" : "morning",
      },
    };
  }

  /** Destino equivale à posição atual (não é mudança). */
  function isSameRelocationTarget(st, iso, t, s) {
    if (st.dateStr !== iso || st.techIndex !== t) return false;
    if (isExtraServiceSlot(st.slotIndex)) {
      const eb = normalizeBooking(st.existing);
      const period =
        eb?.extraPeriod || (st.form.extraPeriod === "afternoon" ? "afternoon" : "morning");
      const wantRow = period === "afternoon" ? AFTERNOON_EXTRA_ROW : MORNING_EXTRA_ROW;
      return s === wantRow;
    }
    const cov = coveredFixedSlotIndices(st.form.timeFrom, st.form.timeTo);
    const clickedStorage = rowToStorageSlot(s);
    const newCov = computeFixedCoverageForRelocation(clickedStorage, cov.length);
    if (!newCov || newCov.length !== cov.length) return false;
    return cov.every((v, i) => v === newCov[i]);
  }

  /**
   * @param {typeof relocationStaging} st
   */
  async function attemptRelocationMove(st, iso, t, s) {
    if (!st) return;
    if (isSameRelocationTarget(st, iso, t, s)) return;

    if (!isCellFreeForRelocation(iso, t, s)) {
      window.alert("Escolha uma célula livre.");
      return;
    }

    const { timeFrom, timeTo } = st.form;
    const origExtra = isExtraServiceSlot(st.slotIndex);

    let canonicalSlotIndex;
    if (origExtra) {
      const eb = normalizeBooking(st.existing);
      const period =
        eb?.extraPeriod || (st.form.extraPeriod === "afternoon" ? "afternoon" : "morning");
      const wantRow = period === "afternoon" ? AFTERNOON_EXTRA_ROW : MORNING_EXTRA_ROW;
      if (s !== wantRow) {
        window.alert(
          `Este serviço extra é de ${period === "morning" ? "manhã" : "tarde"}. Escolha a linha «${
            period === "morning" ? "Extra — manhã" : "Extra — tarde"
          }».`
        );
        return;
      }
      canonicalSlotIndex = wantRow;
    } else {
      if (isExtraServiceSlot(s)) {
        window.alert("Escolha um slot de horário fixo (não a linha extra).");
        return;
      }
      const cov = coveredFixedSlotIndices(timeFrom, timeTo);
      if (cov.length === 0) {
        window.alert("O intervalo não cobre nenhum slot útil (9h–13h ou 14h–18h; pausa de almoço 13h–14h).");
        return;
      }
      const clickedStorage = rowToStorageSlot(s);
      const newCov = computeFixedCoverageForRelocation(clickedStorage, cov.length);
      if (!newCov) {
        window.alert(
          "Não é possível colocar esta marcação a partir deste slot (a duração atravessa a pausa de almoço ou não cabe no horário fixo)."
        );
        return;
      }
      const tr = timeRangeFromStorageIndices(newCov);
      if (!tr) return;
      st.form.timeFrom = tr.timeFrom;
      st.form.timeTo = tr.timeTo;
      canonicalSlotIndex = storageToRow(newCov[0]);
    }

    const techName = getTechDisplayName(t);
    const dateLong = formatDateLongPt(iso);
    const slotLine = slotLabel(canonicalSlotIndex);

    const ok = await showRelocationConfirmDialog(techName, slotLine, dateLong, { isExtra: origExtra });
    if (!ok) return;

    await applyRelocationStagingAndPersist(st, iso, t, canonicalSlotIndex);
  }

  /** Índice da semana (0..n-1) dentro do mês atual. */
  let activeWeekIndex = 0;

  function currentMonthValue() {
    return monthInput.value;
  }

  function parseMonth(val) {
    const [y, m] = val.split("-").map(Number);
    return { year: y, monthIndex: m - 1 };
  }

  function shiftMonth(delta) {
    const { year, monthIndex } = parseMonth(monthInput.value);
    const d = new Date(year, monthIndex + delta, 1);
    monthInput.value = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
  }

  function thisMonthValue() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
  }

  function buildHeadTechnicians(visible) {
    scheduleHead.textContent = "";
    const tr = document.createElement("tr");

    const corner = document.createElement("th");
    corner.colSpan = 2;
    corner.scope = "colgroup";
    corner.className = "grid__corner";
    corner.textContent = "Dia / Horário";
    tr.appendChild(corner);

    visible.forEach(({ name }) => {
      const th = document.createElement("th");
      th.scope = "col";
      th.className = "grid__tech-col";
      th.textContent = name;
      tr.appendChild(th);
    });

    scheduleHead.appendChild(tr);
  }

  function renderWeekTabs(weeks) {
    weekTabs.textContent = "";

    if (weeks.length === 0) return;

    activeWeekIndex = Math.max(0, Math.min(activeWeekIndex, weeks.length - 1));

    weeks.forEach((w, i) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "week-tabs__tab";
      btn.role = "tab";
      btn.setAttribute("aria-selected", i === activeWeekIndex ? "true" : "false");
      btn.id = `week-tab-${i}`;
      btn.setAttribute("aria-controls", "scheduleGrid");
      btn.textContent = `Semana · ${formatShortRange(w.days)}`;
      btn.addEventListener("click", () => {
        activeWeekIndex = i;
        renderWeekTabs(weeks);
        const { year: y, monthIndex: m } = parseMonth(monthInput.value);
        const vis = getVisibleTechnicians();
        renderWeekBody(weeks[activeWeekIndex], y, m, vis);
        scheduleZoneOverlapUiRefresh();
      });
      weekTabs.appendChild(btn);
    });
  }

  function renderWeekBody(week, year, monthIndex, visible) {
    gridBody.textContent = "";
    if (!week || visible.length === 0) return;

    week.days.forEach((dayDate, dayIdx) => {
      const iso = toISODate(dayDate);
      const inMonth = dayDate.getMonth() === monthIndex && dayDate.getFullYear() === year;

      for (let s = 0; s < SLOTS_PER_DAY; s++) {
        const tr = document.createElement("tr");
        if (dayIdx > 0 && s === 0) {
          tr.classList.add("grid__row--day-separator");
        }

        if (s === 0) {
          const thDay = document.createElement("th");
          thDay.rowSpan = SLOTS_PER_DAY;
          thDay.scope = "rowgroup";
          thDay.className = "grid__day-block" + (inMonth ? "" : " grid__day-block--outside");
          const wd = WEEKDAY_SHORT[dayDate.getDay()];
          thDay.innerHTML = `${wd}<span class="grid__day-num">${dayDate.getDate()}</span>`;
          tr.appendChild(thDay);
        }

        const thSlot = document.createElement("th");
        thSlot.scope = "row";
        let slotThClass = "grid__slot-label";
        if (isLunchRow(s)) slotThClass += " grid__slot-label--lunch";
        else if (isMorningExtraRow(s) || isAfternoonExtraRow(s)) slotThClass += " grid__slot-label--extra-banner";
        thSlot.className = slotThClass;
        thSlot.textContent = slotLabel(s);
        tr.appendChild(thSlot);

        visible.forEach(({ index: t }) => {
          const td = document.createElement("td");
          if (isLunchRow(s)) {
            td.className = "grid__lunch-cell";
            td.textContent = inMonth ? "—" : "";
            td.title = inMonth ? "Pausa de almoço (sem marcações)" : "";
            tr.appendChild(td);
            return;
          }

          const cell = getCellDisplayBooking(iso, t, s, inMonth);
          const { booking, isExtra } = cell;
          const btn = document.createElement("button");
          btn.type = "button";
          btn.className = "slot " + (booking ? "slot--busy" : "slot--free");
          if (isExtraServiceSlot(s) && inMonth) btn.classList.add("slot--extra-period");
          if (isExtra && booking && inMonth) btn.classList.add("slot--extra-slot");

          if (!inMonth) {
            btn.classList.add("slot--outside");
            btn.disabled = true;
            btn.textContent = "—";
            btn.title = "Fora do mês selecionado";
          } else {
            btn.dataset.tech = String(t);
            btn.dataset.slot = String(s);
            btn.dataset.date = iso;

            if (booking) {
              applySlotZoneClasses(btn, normalizeBooking(booking).postalCode);
              const multiSeg = getFixedMultiSlotSegment(booking, s, cell.isExtra);
              if (multiSeg && multiSeg.kind === "span") {
                td.classList.add(
                  multiSeg.segment === "first"
                    ? "grid__td--multi-first"
                    : multiSeg.segment === "middle"
                      ? "grid__td--multi-mid"
                      : "grid__td--multi-last",
                );
              }
              fillBookingCell(btn, booking, s, multiSeg);
              const nb = normalizeBooking(booking);
              const stSlot = cell.isExtra ? cell.extraStorage : rowToStorageSlot(s);
              const rid = nb.rangeId ? String(nb.rangeId) : keyFor(iso, t, stSlot);
              btn.dataset.rangeKey = `${iso}|${t}|${rid}`;
              btn.draggable = !relocationPickMode;
              btn.addEventListener("dragstart", (e) => {
                if (relocationPickMode) {
                  e.preventDefault();
                  return;
                }
                dragStaging = buildStagingFromGridBooking(iso, t, s, booking);
                e.dataTransfer.effectAllowed = "move";
                e.dataTransfer.setData("text/plain", "agenda-booking");
                btn.classList.add("slot--dragging");
                document.body.classList.add("agenda-dragging");
              });
              btn.addEventListener("dragend", () => {
                btn.classList.remove("slot--dragging");
                document.body.classList.remove("agenda-dragging");
                document.querySelectorAll(".slot--drop-hover").forEach((el) => el.classList.remove("slot--drop-hover"));
                dragStaging = null;
              });
            } else {
              btn.textContent = "Livre";
              btn.removeAttribute("title");
              delete btn.dataset.rangeKey;
              btn.draggable = false;
            }

            btn.addEventListener("dragover", (e) => {
              if (!dragStaging || relocationPickMode) return;
              if (booking) return;
              e.preventDefault();
              e.dataTransfer.dropEffect = "move";
              btn.classList.add("slot--drop-hover");
            });
            btn.addEventListener("dragleave", (e) => {
              if (!e.relatedTarget || !btn.contains(e.relatedTarget)) btn.classList.remove("slot--drop-hover");
            });
            btn.addEventListener("drop", async (e) => {
              if (!dragStaging || relocationPickMode) return;
              if (booking) return;
              e.preventDefault();
              btn.classList.remove("slot--drop-hover");
              const st = dragStaging;
              dragStaging = null;
              await attemptRelocationMove(st, iso, t, s);
            });

            btn.addEventListener("click", () => {
              if (relocationPickMode) {
                void handleRelocationCellClick(iso, t, s);
                return;
              }
              const cellNow = getCellDisplayBooking(iso, t, s, inMonth);
              if (cellNow.booking) {
                openModal(iso, t, s, cellNow.booking);
              } else {
                openModal(iso, t, s, null);
              }
            });
          }

          td.appendChild(btn);
          tr.appendChild(td);
        });

        gridBody.appendChild(tr);
      }
    });
  }

  /**
   * @param {{ jumpToTodayWeek?: boolean }} [opts]
   */
  function renderGrid(opts) {
    const jumpToTodayWeek = Boolean(opts && opts.jumpToTodayWeek);
    const val = currentMonthValue();
    if (!val) {
      const zoi = document.getElementById("zoneOverlapIndicator");
      if (zoi) zoi.setAttribute("hidden", "");
      clearAllSlotZoneDots();
      return;
    }
    const { year, monthIndex } = parseMonth(val);
    const weeks = weeksOverlappingMonth(year, monthIndex);

    if (weeks.length === 0) {
      scheduleHead.textContent = "";
      gridBody.textContent = "";
      weekTabs.textContent = "";
      tableWrap.hidden = true;
      emptyTechState.hidden = true;
      const zoi = document.getElementById("zoneOverlapIndicator");
      if (zoi) zoi.setAttribute("hidden", "");
      clearAllSlotZoneDots();
      return;
    }

    if (jumpToTodayWeek && val === thisMonthValue()) {
      const idx = weeks.findIndex(weekContainsTodayIso);
      if (idx >= 0) activeWeekIndex = idx;
    } else {
      activeWeekIndex = Math.max(0, Math.min(activeWeekIndex, weeks.length - 1));
    }

    const visible = getVisibleTechnicians();
    if (visible.length === 0) {
      scheduleHead.textContent = "";
      gridBody.textContent = "";
      tableWrap.hidden = true;
      emptyTechState.hidden = false;
      renderWeekTabs(weeks);
      const zoi = document.getElementById("zoneOverlapIndicator");
      if (zoi) zoi.setAttribute("hidden", "");
      clearAllSlotZoneDots();
      return;
    }

    tableWrap.hidden = false;
    emptyTechState.hidden = true;

    buildHeadTechnicians(visible);
    renderWeekTabs(weeks);
    renderWeekBody(weeks[activeWeekIndex], year, monthIndex, visible);
    scheduleZoneOverlapUiRefresh();
  }

  /** @param {{ extraFromButton?: boolean, wizardPostal?: string, wizardGas?: boolean }} [opts] */
  function openModal(dateStr, techIndex, slotIndex, existing, opts) {
    const extraFromButton = opts && opts.extraFromButton;
    const wizardPostal = opts && opts.wizardPostal;
    const wizardGas = opts && opts.wizardGas;
    editContext = {
      dateStr,
      techIndex,
      slotIndex,
      existing,
      extraFromButton: Boolean(extraFromButton),
      gasFromWizard: Boolean(wizardGas),
    };
    if (extraMetaFields) extraMetaFields.hidden = !extraFromButton;
    if (extraFromButton && fieldExtraDate && fieldExtraTech) {
      fieldExtraDate.min = todayISODate();
      fieldExtraDate.value = dateStr;
      fieldExtraTech.innerHTML = "";
      getVisibleTechnicians().forEach(({ index, name }) => {
        const opt = document.createElement("option");
        opt.value = String(index);
        opt.textContent = name;
        fieldExtraTech.appendChild(opt);
      });
      fieldExtraTech.value = String(techIndex);
      setExtraPeriodRadio(slotIndex === AFTERNOON_EXTRA_ROW ? "afternoon" : "morning");
      syncExtraSlotFromPeriodForButton();
    }
    const techName = getTechDisplayName(techIndex);
    const when = new Date(dateStr + "T12:00:00");
    const dateFmt = when.toLocaleDateString("pt-PT", {
      weekday: "long",
      day: "numeric",
      month: "long",
      year: "numeric",
    });

    const extra = isExtraServiceSlot(slotIndex);
    timeRangeFields.hidden = extra;
    if (fieldSlotSpanGroup) fieldSlotSpanGroup.hidden = extra;
    setTimeFieldsRequired(!extra);
    if (timeRangeHint) {
      if (extra) {
        timeRangeHint.textContent = extraFromButton
          ? "O período (manhã ou tarde) define a linha na grelha; não é necessário horário de início/fim."
          : "Serviço extra nesta linha; não é necessário horário de início/fim.";
      } else {
        timeRangeHint.textContent =
          "Escolha simples (1 h), duplo, triplo ou quádruplo; ou ajuste início/fim. O bloco não atravessa a pausa 13h–14h.";
      }
    }

    const ex = normalizeBooking(existing);
    if (btnRelocateBooking) {
      btnRelocateBooking.hidden = !(existing && !extraFromButton);
    }
    if (extraFromButton) {
      modalTitle.textContent = "Serviço extra";
      modalContext.textContent = `${techName} · Serviço extra · ${dateFmt}`;
    } else {
      modalContext.textContent = `${techName} · ${slotLabel(slotIndex)} · ${dateFmt}`;
    }

    if (extra) {
      fieldTimeFrom.value = "";
      fieldTimeTo.value = "";
      if (extraFromButton) {
        setExtraPeriodRadio(slotIndex === AFTERNOON_EXTRA_ROW ? "afternoon" : "morning");
        syncExtraSlotFromPeriodForButton();
      } else if (existing) {
        setExtraPeriodRadio(ex.extraPeriod === "afternoon" ? "afternoon" : "morning");
      } else {
        setExtraPeriodRadio(slotIndex === AFTERNOON_EXTRA_ROW ? "afternoon" : "morning");
      }
    }

    if (existing && !extraFromButton) {
      modalTitle.textContent = "Marcação";
      fieldProcess.value = ex.processNo;
      fieldClientName.value = ex.clientName;
      fieldPhone.value = ex.phone;
      fieldPostal.value = ex.postalCode;
      fieldDevice.value = ex.device;
      fieldObservations.value = ex.observations;
      btnDelete.hidden = false;
      if (!extra) {
        if (ex.timeFrom && ex.timeTo) {
          fieldTimeFrom.value = ex.timeFrom;
          fieldTimeTo.value = ex.timeTo;
        } else {
          fieldTimeFrom.value = hhmmSlotStartForRow(slotIndex);
          fieldTimeTo.value = hhmmSlotEndForRow(slotIndex);
        }
        const cov = coveredFixedSlotIndices(fieldTimeFrom.value, fieldTimeTo.value);
        setSlotSpanRadios(cov.length > 0 ? cov.length : 1);
        syncSlotSpanRadiosFromTimeFields();
      }
    } else if (!existing) {
      modalTitle.textContent = extraFromButton ? "Serviço extra" : "Nova marcação";
      fieldProcess.value = "";
      fieldClientName.value = "";
      fieldPhone.value = "";
      fieldPostal.value = wizardPostal ? String(wizardPostal).trim() : "";
      fieldDevice.value = "";
      fieldObservations.value = "";
      btnDelete.hidden = true;
      if (!extra) {
        fieldTimeFrom.value = hhmmSlotStartForRow(slotIndex);
        fieldTimeTo.value = hhmmSlotEndForRow(slotIndex);
        setSlotSpanRadios(1);
        syncSlotSpanRadiosFromTimeFields();
      }
    }

    applyPostalZoneClassesToInput(fieldPostal);
    modal.showModal();
    if (extraFromButton && fieldExtraDate) {
      fieldExtraDate.focus();
    } else if (!extra) {
      fieldProcess.focus();
    } else {
      fieldProcess.focus();
    }
  }

  function closeModal() {
    editContext = null;
    if (extraMetaFields) extraMetaFields.hidden = true;
    if (btnRelocateBooking) btnRelocateBooking.hidden = true;
    modal.close();
  }

  function captureRelocationStaging() {
    if (!editContext) return null;
    return {
      dateStr: editContext.dateStr,
      techIndex: editContext.techIndex,
      slotIndex: editContext.slotIndex,
      existing: editContext.existing,
      form: {
        processNo: fieldProcess.value,
        clientName: fieldClientName.value,
        phone: fieldPhone.value,
        postalCode: fieldPostal.value,
        device: fieldDevice.value,
        observations: fieldObservations.value,
        timeFrom: fieldTimeFrom.value,
        timeTo: fieldTimeTo.value,
        extraPeriod:
          editContext && isExtraServiceSlot(editContext.slotIndex)
            ? editContext.slotIndex === MORNING_EXTRA_ROW
              ? "morning"
              : "afternoon"
            : getExtraPeriodFromRadio(),
      },
      extraFromButton: Boolean(editContext.extraFromButton),
    };
  }

  function restoreFormFromStaging(st) {
    const f = st.form;
    fieldProcess.value = f.processNo;
    fieldClientName.value = f.clientName;
    fieldPhone.value = f.phone;
    fieldPostal.value = f.postalCode;
    fieldDevice.value = f.device;
    fieldObservations.value = f.observations;
    fieldTimeFrom.value = f.timeFrom;
    fieldTimeTo.value = f.timeTo;
    if (f.extraPeriod === "afternoon") setExtraPeriodRadio("afternoon");
    else if (f.extraPeriod === "morning") setExtraPeriodRadio("morning");
    syncSlotSpanRadiosFromTimeFields();
    applyPostalZoneClassesToInput(fieldPostal);
  }

  function setRelocationBannerVisible(visible) {
    const b = document.getElementById("relocationBanner");
    if (b) b.hidden = !visible;
  }

  /** @returns {Promise<boolean>} */
  function showRelocationConfirmDialog(techName, slotLine, dateLong, opts) {
    const isExtra = Boolean(opts && opts.isExtra);
    const dlg = document.getElementById("relocationConfirmDialog");
    const textEl = document.getElementById("relocationConfirmText");
    const btnOk = document.getElementById("btnRelocationConfirmOk");
    const btnCancel = document.getElementById("btnRelocationConfirmCancel");
    const tail = isExtra
      ? "O período (extra) e os dados do cliente mantêm-se."
      : "O intervalo de horário ajusta-se à célula escolhida (mesma duração). Os dados do cliente mantêm-se.";
    if (!dlg || !textEl || !btnOk || !btnCancel) {
      return Promise.resolve(
        window.confirm(`Confirma registar a marcação em ${techName} · ${slotLine} · ${dateLong}? ${tail}`)
      );
    }
    textEl.textContent = `Confirma registar a marcação em ${techName} · ${slotLine} · ${dateLong}? ${tail}`;
    return new Promise((resolve) => {
      const cleanup = () => {
        btnOk.removeEventListener("click", onOk);
        btnCancel.removeEventListener("click", onCancel);
        dlg.removeEventListener("cancel", onEsc);
      };
      const onOk = () => {
        cleanup();
        dlg.close();
        resolve(true);
      };
      const onCancel = () => {
        cleanup();
        dlg.close();
        resolve(false);
      };
      const onEsc = () => {
        cleanup();
        resolve(false);
      };
      btnOk.addEventListener("click", onOk);
      btnCancel.addEventListener("click", onCancel);
      dlg.addEventListener("cancel", onEsc);
      dlg.showModal();
    });
  }

  function beginRelocationPicker() {
    if (!editContext) return;
    const snap = captureRelocationStaging();
    if (!snap) return;
    relocationStaging = snap;
    relocationPickMode = true;
    if (extraMetaFields) extraMetaFields.hidden = true;
    modal.close();
    editContext = null;
    document.body.classList.add("relocation-pick-mode");
    setRelocationBannerVisible(true);

    const ym = snap.dateStr.slice(0, 7);
    if (monthInput.value !== ym) {
      monthInput.value = ym;
    }
    const { year, monthIndex } = parseMonth(monthInput.value);
    const weeks = weeksOverlappingMonth(year, monthIndex);
    const widx = weeks.findIndex((w) => weekContainsIsoDate(w, snap.dateStr));
    activeWeekIndex = widx >= 0 ? widx : 0;
    renderGrid({});
  }

  function cancelRelocationPick() {
    const snap = relocationStaging;
    relocationStaging = null;
    relocationPickMode = false;
    document.body.classList.remove("relocation-pick-mode");
    setRelocationBannerVisible(false);
    if (!snap) return;
    if (snap.extraFromButton && extraMetaFields && fieldExtraDate && fieldExtraTech) {
      extraMetaFields.hidden = false;
      fieldExtraDate.value = snap.dateStr;
      fieldExtraTech.value = String(snap.techIndex);
    } else if (extraMetaFields) {
      extraMetaFields.hidden = true;
    }
    openModal(snap.dateStr, snap.techIndex, snap.slotIndex, snap.existing, {
      extraFromButton: Boolean(snap.extraFromButton),
    });
    restoreFormFromStaging(snap);
    if (snap.extraFromButton) syncExtraSlotFromPeriodForButton();
  }

  async function applyRelocationStagingAndPersist(st, dateStr, techIndex, slotIndex) {
    if (!st) return;

    const prevB = normalizeBooking(st.existing);
    const postalForZone = String(st.form.postalCode || "").trim();
    if (!isValidPostalPt(postalForZone)) {
      window.alert("Indique um código postal português válido (1234-567).");
      return;
    }
    const zoneResult = await alertNearbyZoneIfNeeded(postalForZone, dateStr, techIndex, prevB?.rangeId);
    if (zoneResult === "cancel") return;
    if (zoneResult === "change") {
      editContext = {
        dateStr: st.dateStr,
        techIndex: st.techIndex,
        slotIndex: st.slotIndex,
        existing: st.existing,
        extraFromButton: Boolean(st.extraFromButton),
      };
      restoreFormFromStaging(st);
      if (extraMetaFields) extraMetaFields.hidden = true;
      beginRelocationPicker();
      return;
    }

    if (st.existing) {
      const b = normalizeBooking(st.existing);
      if (b?.rangeId) {
        removeRangeFromStorage(st.dateStr, st.techIndex, b.rangeId);
      } else if (isExtraServiceSlot(st.slotIndex)) {
        const oldSt = b.extraPeriod === "afternoon" ? 9 : 8;
        setBookingAt(st.dateStr, st.techIndex, oldSt, null);
      } else {
        setBookingAt(st.dateStr, st.techIndex, rowToStorageSlot(st.slotIndex), null);
      }
    }

    relocationStaging = null;
    relocationPickMode = false;
    document.body.classList.remove("relocation-pick-mode");
    setRelocationBannerVisible(false);

    editContext = {
      dateStr,
      techIndex,
      slotIndex,
      existing: st.existing,
      extraFromButton: Boolean(st.extraFromButton),
      skipRemoveExisting: true,
    };
    restoreFormFromStaging(st);
    if (st.extraFromButton && extraMetaFields && fieldExtraDate && fieldExtraTech) {
      extraMetaFields.hidden = false;
      fieldExtraDate.value = dateStr;
      fieldExtraTech.value = String(techIndex);
    }

    if (!st.extraFromButton && extraMetaFields) extraMetaFields.hidden = true;

    const techName = getTechDisplayName(techIndex);
    const when = new Date(dateStr + "T12:00:00");
    const dateFmt = when.toLocaleDateString("pt-PT", {
      weekday: "long",
      day: "numeric",
      month: "long",
      year: "numeric",
    });

    const ex = normalizeBooking(st.existing);
    const extra = isExtraServiceSlot(slotIndex);
    if (st.extraFromButton) {
      modalTitle.textContent = "Serviço extra";
      modalContext.textContent = `${techName} · Serviço extra · ${dateFmt}`;
    } else {
      modalContext.textContent = `${techName} · ${slotLabel(slotIndex)} · ${dateFmt}`;
      modalTitle.textContent = ex ? "Marcação" : "Nova marcação";
    }
    timeRangeFields.hidden = extra;
    if (fieldSlotSpanGroup) fieldSlotSpanGroup.hidden = extra;
    setTimeFieldsRequired(!extra);
    if (timeRangeHint) {
      if (extra) {
        timeRangeHint.textContent = st.extraFromButton
          ? "O período (manhã ou tarde) define a linha na grelha; não é necessário horário de início/fim."
          : "Serviço extra nesta linha; não é necessário horário de início/fim.";
      } else {
        timeRangeHint.textContent =
          "Escolha simples (1 h), duplo, triplo ou quádruplo; ou ajuste início/fim. O bloco não atravessa a pausa 13h–14h.";
      }
    }
    btnDelete.hidden = !st.existing;
    if (btnRelocateBooking) btnRelocateBooking.hidden = !(st.existing && !st.extraFromButton);

    modal.showModal();
    await executePersistBooking(true);
  }

  async function handleRelocationCellClick(iso, t, s) {
    if (!relocationStaging) return;
    await attemptRelocationMove(relocationStaging, iso, t, s);
  }

  const btnRelocationPickCancel = document.getElementById("btnRelocationPickCancel");
  if (btnRelocationPickCancel) btnRelocationPickCancel.addEventListener("click", cancelRelocationPick);

  if (btnRelocateBooking) {
    btnRelocateBooking.addEventListener("click", () => {
      if (!editContext || !editContext.existing) return;
      beginRelocationPicker();
    });
  }

  modalForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    await executePersistBooking(false);
  });

  /**
   * @param {boolean} skipZoneCheck — após «Alterar» no aviso de zona, grava sem voltar a perguntar
   */
  async function executePersistBooking(skipZoneCheck) {
    if (!editContext) return;

    const processNo = fieldProcess.value.trim();
    const clientName = fieldClientName.value.trim();
    const phone = fieldPhone.value.trim();
    const postalCode = fieldPostal.value.trim();
    const device = fieldDevice.value.trim();
    const observations = fieldObservations.value.trim();

    if (!processNo || !clientName || !phone || !device) {
      window.alert("Preencha número de processo, cliente, telefone e aparelho (campos obrigatórios).");
      if (!processNo) fieldProcess.focus();
      else if (!clientName) fieldClientName.focus();
      else if (!phone) fieldPhone.focus();
      else fieldDevice.focus();
      return;
    }

    if (!isValidPostalPt(postalCode)) {
      alert("Indique um código postal português válido (1234-567).");
      fieldPostal.focus();
      return;
    }

    if (editContext.extraFromButton && fieldExtraDate && fieldExtraTech) {
      const d = fieldExtraDate.value.trim();
      const ti = Number(fieldExtraTech.value);
      if (!d || !Number.isFinite(ti)) {
        alert("Indique o dia e o técnico.");
        fieldExtraDate.focus();
        return;
      }
      editContext.dateStr = d;
      editContext.techIndex = ti;
      syncExtraSlotFromPeriodForButton();
    }

    let prev = normalizeBooking(editContext.existing);
    let dateStr = editContext.dateStr;
    let techIndex = editContext.techIndex;
    let slotIndex = editContext.slotIndex;

    if (!skipZoneCheck) {
      const zoneResult = await alertNearbyZoneIfNeeded(postalCode, dateStr, techIndex, prev?.rangeId);
      if (zoneResult === "cancel") return;
      if (zoneResult === "change") {
        beginRelocationPicker();
        return;
      }
    }

    prev = normalizeBooking(editContext.existing);
    const extra = isExtraServiceSlot(editContext.slotIndex);
    dateStr = editContext.dateStr;
    techIndex = editContext.techIndex;
    slotIndex = editContext.slotIndex;
    const storageSlot = rowToStorageSlot(slotIndex);

    let timeFrom = "";
    let timeTo = "";
    /** @type {number[]} */
    let covered = [];

    if (!extra) {
      const anchorSt = getCoverageAnchorStorageForModal();
      const span = getSlotSpanFromRadios();
      const cov = computeFixedCoverageForRelocation(anchorSt, span);
      if (!cov || cov.length === 0) {
        alert("A duração escolhida não cabe neste intervalo (verifique a pausa 13h–14h).");
        return;
      }
      const tr = timeRangeFromStorageIndices(cov);
      if (!tr) {
        alert("Não foi possível calcular o horário.");
        return;
      }
      timeFrom = tr.timeFrom;
      timeTo = tr.timeTo;
      covered = cov;
      if (fieldTimeFrom) fieldTimeFrom.value = timeFrom;
      if (fieldTimeTo) fieldTimeTo.value = timeTo;
    }

    const rangeId = prev?.rangeId || newRangeId();

    /** @type {Record<string, unknown>} */
    const payload = {
      processNo,
      clientName,
      phone,
      postalCode,
      device,
      observations,
      createdAt: prev?.createdAt ?? Date.now(),
      rangeId,
      gasAppliance: editContext.gasFromWizard ? true : Boolean(prev?.gasAppliance),
    };

    if (extra) {
      const extraPeriod = slotIndex === MORNING_EXTRA_ROW ? "morning" : "afternoon";
      const targetStorage = extraPeriod === "morning" ? 8 : 9;
      payload.extraPeriod = extraPeriod;

      if (!editContext.skipRemoveExisting && editContext.existing) {
        if (prev?.rangeId) {
          removeRangeFromStorage(dateStr, techIndex, prev.rangeId);
        } else {
          const oldB = normalizeBooking(editContext.existing);
          const oldSt = oldB.extraPeriod === "afternoon" ? 9 : 8;
          setBookingAt(dateStr, techIndex, oldSt, null);
        }
      }

      if (getBookingAt(dateStr, techIndex, targetStorage)) {
        alert(
          targetStorage === 8
            ? "Já existe um serviço extra de manhã neste dia para este técnico."
            : "Já existe um serviço extra de tarde neste dia para este técnico."
        );
        renderGrid({});
        return;
      }
      setBookingAt(dateStr, techIndex, targetStorage, payload);
      if (editContext.skipRemoveExisting) delete editContext.skipRemoveExisting;
      closeModal();
      renderGrid({});
      return;
    }

    Object.assign(payload, { timeFrom, timeTo });

    const ignoreRid = prev?.rangeId;
    const legacyEdit = Boolean(editContext.existing && !ignoreRid);

    if (
      wouldConflict(dateStr, techIndex, covered, {
        ignoreRangeId: ignoreRid,
        legacyEditSlot: legacyEdit && !extra ? storageSlot : null,
      })
    ) {
      alert("Um ou mais destes slots já têm outra marcação. Ajuste o horário ou liberte o espaço.");
      return;
    }

    if (editContext.existing && !editContext.skipRemoveExisting) {
      if (prev?.rangeId) removeRangeFromStorage(dateStr, techIndex, prev.rangeId);
      else setBookingAt(dateStr, techIndex, storageSlot, null);
    }
    if (editContext.skipRemoveExisting) delete editContext.skipRemoveExisting;

    setBookingsAcrossSlots(dateStr, techIndex, covered, payload);
    closeModal();
    renderGrid({});
  }

  btnDelete.addEventListener("click", () => {
    if (!editContext || !editContext.existing) return;
    if (!confirm("Cancelar esta marcação?")) return;
    const b = normalizeBooking(editContext.existing);
    const { dateStr, techIndex, slotIndex } = editContext;
    if (b?.rangeId) {
      removeRangeFromStorage(dateStr, techIndex, b.rangeId);
    } else if (isExtraServiceSlot(slotIndex)) {
      const st = b.extraPeriod === "afternoon" ? 9 : 8;
      setBookingAt(dateStr, techIndex, st, null);
    } else {
      setBookingAt(dateStr, techIndex, rowToStorageSlot(slotIndex), null);
    }
    closeModal();
    renderGrid({});
  });

  btnClose.addEventListener("click", closeModal);

  monthInput.addEventListener("change", () => {
    activeWeekIndex = 0;
    renderGrid({ jumpToTodayWeek: monthInput.value === thisMonthValue() });
  });
  btnPrevMonth.addEventListener("click", () => {
    shiftMonth(-1);
    activeWeekIndex = 0;
    renderGrid({ jumpToTodayWeek: monthInput.value === thisMonthValue() });
  });
  btnNextMonth.addEventListener("click", () => {
    shiftMonth(1);
    activeWeekIndex = 0;
    renderGrid({ jumpToTodayWeek: monthInput.value === thisMonthValue() });
  });
  btnToday.addEventListener("click", () => {
    monthInput.value = thisMonthValue();
    activeWeekIndex = 0;
    renderGrid({ jumpToTodayWeek: true });
  });

  async function runWizardSearch() {
    if (!fieldWizardPostal || !wizardTechResults || !btnWizardContinue) return;
    const cp = fieldWizardPostal.value.trim();
    if (!isValidPostalPt(cp)) {
      window.alert("Indique um código postal português válido (1234-567).");
      fieldWizardPostal.focus();
      return;
    }
    const isGas = fieldWizardGas && fieldWizardGas.checked;
    wizardSuggest = null;
    btnWizardContinue.hidden = true;
    wizardTechResults.hidden = false;
    wizardTechResults.innerHTML = '<p class="wizard-tech-results__loading">A calcular…</p>';

    const vis = getVisibleTechnicians();
    if (vis.length === 0) {
      wizardTechResults.innerHTML =
        '<p class="wizard-tech-results__error">Configure pelo menos um técnico em «Técnicos».</p>';
      return;
    }

    try {
    if (isGas) {
      const { names } = loadConfig();
      const authorizedIndices = [];
      for (let i = 0; i < TECH_COUNT; i++) {
        if (isGasAuthorizedTechName(names[i])) authorizedIndices.push(i);
      }
      if (authorizedIndices.length === 0) {
        wizardTechResults.innerHTML =
          '<p class="wizard-tech-results__error">Para aparelhos a gás é necessário um técnico com o nome <strong>João Rocha</strong> ou <strong>Marcos Correia</strong> numa das posições em «Técnicos».</p>';
        return;
      }
      const authSet = new Set(authorizedIndices);
      const visAuth = vis.filter((v) => authSet.has(v.index));
      if (visAuth.length === 0) {
        wizardTechResults.innerHTML =
          '<p class="wizard-tech-results__error">Nenhum técnico autorizado para gás está visível na grelha. Atribua o nome em «Técnicos».</p>';
        return;
      }

      const zoneGas = await findBestZoneWizardMatch(cp, todayISODate(), visAuth.map((v) => v.index));
      if (zoneGas) {
        const pickName = getTechDisplayName(zoneGas.techIndex);
        const counts = formatWizardSlotCountsHtml(zoneGas.dateStr, zoneGas.techIndex);
        const dist = zoneGas.km.toFixed(1);
        wizardSuggest = {
          dateStr: zoneGas.dateStr,
          techIndex: zoneGas.techIndex,
          slotIndex: zoneGas.slotIndex,
          isGas: true,
        };
        if (zoneGas.kind === "zone_fixed") {
          wizardTechResults.innerHTML = `<p class="wizard-tech-results__ok">Trabalho a gás: há marcações na zona (≤${ZONE_RADIUS_KM} km, ≈${dist} km ao CP mais próximo). <strong>${pickName}</strong> — sugestão em horário normal: <strong>${formatDateLongPt(zoneGas.dateStr)}</strong> · ${slotLabel(zoneGas.slotIndex)}. ${counts}</p>`;
        } else {
          wizardTechResults.innerHTML = `<p class="wizard-tech-results__warn">Trabalho a gás: <strong>${pickName}</strong> está na zona (≤${ZONE_RADIUS_KM} km, ≈${dist} km) mas só há disponibilidade em <strong>serviço extra</strong> (${zoneGas.period === "morning" ? "manhã" : "tarde"}). ${counts}</p>`;
        }
        btnWizardContinue.hidden = false;
        return;
      }

      const ranked = await rankTechniciansByDistanceFromCustomerPostal(cp);
      const rankedAuth = ranked.filter((t) => authSet.has(t.index));
      let pick;
      let rankedHadDistance = false;
      if (rankedAuth.length > 0) {
        pick = rankedAuth[0];
        rankedHadDistance = true;
      } else {
        visAuth.sort(
          (a, b) => gasAuthorizedFallbackRank(a.name) - gasAuthorizedFallbackRank(b.name) || a.index - b.index,
        );
        pick = visAuth[0];
      }

      const distText = rankedHadDistance ? pick.km.toFixed(1) : "—";
      if (rankedHadDistance && pick.km > NEW_ROUTE_DISTANCE_KM) {
        const next = findNextFreeExtraDay(pick.index, todayISODate());
        if (!next) {
          wizardTechResults.innerHTML = `<p class="wizard-tech-results__error">Não foi encontrado slot extra livre para <strong>${pick.name}</strong> nos próximos dias.</p>`;
          return;
        }
        wizardSuggest = {
          dateStr: next.dateStr,
          techIndex: pick.index,
          slotIndex: next.slotIndex,
          isGas: true,
        };
        const counts = formatWizardSlotCountsHtml(next.dateStr, pick.index);
        wizardTechResults.innerHTML = `<p class="wizard-tech-results__warn">Trabalho a gás: sem marcações na zona. Técnico autorizado mais próximo: <strong>${pick.name}</strong> (${distText} km). Como está a mais de ${NEW_ROUTE_DISTANCE_KM} km, sugere-se <strong>nova rota</strong> noutro dia: <strong>${formatDateLongPt(next.dateStr)}</strong> (${next.period === "morning" ? "extra manhã" : "extra tarde"}). ${counts}</p>`;
        btnWizardContinue.hidden = false;
        return;
      }

      const slotGas = findFirstAvailableSlotPreferringFixed(pick.index, todayISODate());
      if (!slotGas) {
        wizardTechResults.innerHTML = `<p class="wizard-tech-results__error">Não foi encontrada vaga livre para <strong>${pick.name}</strong> nos próximos dias.</p>`;
        return;
      }
      wizardSuggest = {
        dateStr: slotGas.dateStr,
        techIndex: pick.index,
        slotIndex: slotGas.slotIndex,
        isGas: true,
      };
      const counts = formatWizardSlotCountsHtml(slotGas.dateStr, pick.index);
      const label = slotLabel(slotGas.slotIndex);
      const extraNote = slotGas.isExtra
        ? ` <strong>Atenção:</strong> só há vaga em serviço extra (${slotGas.period === "morning" ? "manhã" : "tarde"}).`
        : "";
      if (rankedHadDistance) {
        wizardTechResults.innerHTML = `<p class="wizard-tech-results__ok">Trabalho a gás: sem marcações na zona. Técnico autorizado mais próximo: <strong>${pick.name}</strong> (${distText} km). Sugestão: <strong>${formatDateLongPt(slotGas.dateStr)}</strong> · ${label}. ${counts}${extraNote}</p>`;
      } else {
        wizardTechResults.innerHTML = `<p class="wizard-tech-results__warn">Trabalho a gás: <strong>${pick.name}</strong> (sem CPs na agenda nem CP de referência para distância). Sugestão: <strong>${formatDateLongPt(slotGas.dateStr)}</strong> · ${label}. ${counts}${extraNote}</p>`;
      }
      btnWizardContinue.hidden = false;
      return;
    }

    const zone = await findBestZoneWizardMatch(cp, todayISODate(), vis.map((v) => v.index));
    if (zone) {
      const pickName = getTechDisplayName(zone.techIndex);
      const counts = formatWizardSlotCountsHtml(zone.dateStr, zone.techIndex);
      const dist = zone.km.toFixed(1);
      wizardSuggest = {
        dateStr: zone.dateStr,
        techIndex: zone.techIndex,
        slotIndex: zone.slotIndex,
        isGas: false,
      };
      if (zone.kind === "zone_fixed") {
        wizardTechResults.innerHTML = `<p class="wizard-tech-results__ok">Há marcações na zona (≤${ZONE_RADIUS_KM} km, ≈${dist} km ao CP mais próximo). <strong>${pickName}</strong> — sugestão em horário normal: <strong>${formatDateLongPt(zone.dateStr)}</strong> · ${slotLabel(zone.slotIndex)}. ${counts}</p>`;
      } else {
        wizardTechResults.innerHTML = `<p class="wizard-tech-results__warn"><strong>${pickName}</strong> está na zona (≤${ZONE_RADIUS_KM} km, ≈${dist} km) mas só há disponibilidade em <strong>serviço extra</strong> (${zone.period === "morning" ? "manhã" : "tarde"}). ${counts}</p>`;
      }
      btnWizardContinue.hidden = false;
      return;
    }

    const ranked = await rankTechniciansByDistanceFromCustomerPostal(cp);
    const visSet = new Set(vis.map((v) => v.index));
    const rankedVis = ranked.filter((t) => visSet.has(t.index));

    if (rankedVis.length === 0) {
      const pick = vis[0];
      const slot = findFirstAvailableSlotPreferringFixed(pick.index, todayISODate());
      if (!slot) {
        wizardTechResults.innerHTML =
          '<p class="wizard-tech-results__error">Sem vagas livres. Adicione marcações com CP na agenda ou configure CP de referência em «Técnicos».</p>';
        return;
      }
      wizardSuggest = {
        dateStr: slot.dateStr,
        techIndex: pick.index,
        slotIndex: slot.slotIndex,
        isGas: false,
      };
      const counts = formatWizardSlotCountsHtml(slot.dateStr, pick.index);
      const label = slotLabel(slot.slotIndex);
      const extraNote = slot.isExtra
        ? ` <strong>Atenção:</strong> só há vaga em serviço extra (${slot.period === "morning" ? "manhã" : "tarde"}).`
        : "";
      wizardTechResults.innerHTML = `<p class="wizard-tech-results__warn">Não há CPs nas marcações nem CP de referência para calcular distâncias. Sugestão (1.º técnico visível): <strong>${pick.name}</strong> — <strong>${formatDateLongPt(slot.dateStr)}</strong> · ${label}. ${counts}${extraNote}</p>`;
      btnWizardContinue.hidden = false;
      return;
    }

    const pick = rankedVis[0];
    const distText = pick.km.toFixed(1);

    if (pick.km > NEW_ROUTE_DISTANCE_KM) {
      const next = findNextFreeExtraDay(pick.index, todayISODate());
      if (!next) {
        wizardTechResults.innerHTML =
          '<p class="wizard-tech-results__error">Sem slots extra livres para o técnico sugerido.</p>';
        return;
      }
      wizardSuggest = {
        dateStr: next.dateStr,
        techIndex: pick.index,
        slotIndex: next.slotIndex,
        isGas: false,
      };
      const counts = formatWizardSlotCountsHtml(next.dateStr, pick.index);
      wizardTechResults.innerHTML = `<p class="wizard-tech-results__warn">Sem marcações na zona. Técnico mais próximo: <strong>${pick.name}</strong> (${distText} km). Como está a mais de ${NEW_ROUTE_DISTANCE_KM} km, sugere-se <strong>nova rota</strong> noutro dia: <strong>${formatDateLongPt(next.dateStr)}</strong> (${next.period === "morning" ? "extra manhã" : "extra tarde"}). ${counts}</p>`;
      btnWizardContinue.hidden = false;
      return;
    }

    const slot = findFirstAvailableSlotPreferringFixed(pick.index, todayISODate());
    if (!slot) {
      wizardTechResults.innerHTML =
        '<p class="wizard-tech-results__error">Não foi encontrada vaga livre para o técnico sugerido nos próximos dias.</p>';
      return;
    }
    wizardSuggest = {
      dateStr: slot.dateStr,
      techIndex: pick.index,
      slotIndex: slot.slotIndex,
      isGas: false,
    };
    const counts = formatWizardSlotCountsHtml(slot.dateStr, pick.index);
    const label = slotLabel(slot.slotIndex);
    const extraNote = slot.isExtra
      ? ` <strong>Atenção:</strong> só há vaga em serviço extra (${slot.period === "morning" ? "manhã" : "tarde"}).`
      : "";
    const cls = pick.km <= ZONE_RADIUS_KM ? "wizard-tech-results__ok" : "wizard-tech-results__warn";
    wizardTechResults.innerHTML = `<p class="${cls}">Sem marcações na zona (≤${ZONE_RADIUS_KM} km). Técnico mais próximo: <strong>${pick.name}</strong> (${distText} km). Sugestão: <strong>${formatDateLongPt(slot.dateStr)}</strong> · ${label}. ${counts}${extraNote}</p>`;
    btnWizardContinue.hidden = false;
    } catch (err) {
      console.error("runWizardSearch", err);
      const msg = err && err.message ? String(err.message) : String(err);
      const safe = msg.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
      wizardTechResults.innerHTML = `<p class="wizard-tech-results__error">Não foi possível concluir a pesquisa. ${safe} Se não usa Google Maps, desative a chave em «Técnicos» para usar só geocodificação gratuita (OpenStreetMap).</p>`;
    }
  }

  function openWizardSuggestedBooking() {
    if (!wizardSuggest || !fieldWizardPostal) return;
    const { dateStr, techIndex, slotIndex } = wizardSuggest;
    const ym = dateStr.slice(0, 7);
    if (monthInput.value !== ym) {
      monthInput.value = ym;
    }
    const { year, monthIndex } = parseMonth(monthInput.value);
    const weeks = weeksOverlappingMonth(year, monthIndex);
    const widx = weeks.findIndex((w) => weekContainsIsoDate(w, dateStr));
    activeWeekIndex = widx >= 0 ? widx : 0;
    renderGrid({});
    const cp = fieldWizardPostal.value.trim();
    const gas = fieldWizardGas && fieldWizardGas.checked;
    if (dialogScheduleByPostal) dialogScheduleByPostal.close();
    openModal(dateStr, techIndex, slotIndex, null, {
      extraFromButton: isExtraServiceSlot(slotIndex),
      wizardPostal: cp,
      wizardGas: gas,
    });
  }

  if (btnScheduleByPostal && dialogScheduleByPostal) {
    btnScheduleByPostal.addEventListener("click", () => {
      if (getVisibleTechnicians().length === 0) {
        window.alert("Configure pelo menos um técnico em «Técnicos».");
        return;
      }
      wizardSuggest = null;
      if (fieldWizardPostal) fieldWizardPostal.value = "";
      if (fieldWizardGas) fieldWizardGas.checked = false;
      if (wizardTechResults) {
        wizardTechResults.hidden = true;
        wizardTechResults.textContent = "";
      }
      if (btnWizardContinue) btnWizardContinue.hidden = true;
      dialogScheduleByPostal.showModal();
      if (fieldWizardPostal) fieldWizardPostal.focus();
    });
  }
  if (btnWizardClose) {
    btnWizardClose.addEventListener("click", () => dialogScheduleByPostal && dialogScheduleByPostal.close());
  }
  if (btnWizardSearch) {
    btnWizardSearch.addEventListener("click", () => void runWizardSearch());
  }
  if (btnWizardContinue) {
    btnWizardContinue.addEventListener("click", openWizardSuggestedBooking);
  }

  const formScheduleByPostal = document.getElementById("formScheduleByPostal");
  if (formScheduleByPostal) {
    formScheduleByPostal.addEventListener("submit", (e) => e.preventDefault());
  }

  if (fieldExtraPeriodMorning) {
    fieldExtraPeriodMorning.addEventListener("change", syncExtraSlotFromPeriodForButton);
  }
  if (fieldExtraPeriodAfternoon) {
    fieldExtraPeriodAfternoon.addEventListener("change", syncExtraSlotFromPeriodForButton);
  }

  function ensureTechNameFields() {
    if (techNamesFields.children.length && !document.getElementById("techBaseCp0")) {
      techNamesFields.textContent = "";
    }
    if (techNamesFields.children.length) return;
    for (let i = 0; i < TECH_COUNT; i++) {
      const wrap = document.createElement("div");
      wrap.className = "tech-names__row";
      const nameLab = document.createElement("label");
      nameLab.className = "field tech-names__field";
      const sp1 = document.createElement("span");
      sp1.className = "field__label";
      sp1.textContent = `Posição ${i + 1} — nome`;
      const inpName = document.createElement("input");
      inpName.type = "text";
      inpName.id = `techName${i}`;
      inpName.autocomplete = "name";
      inpName.placeholder = "Nome (vazio = coluna oculta)";
      nameLab.appendChild(sp1);
      nameLab.appendChild(inpName);

      const cpLab = document.createElement("label");
      cpLab.className = "field tech-names__field";
      const sp2 = document.createElement("span");
      sp2.className = "field__label";
      sp2.textContent = "CP referência (opcional)";
      const inpCp = document.createElement("input");
      inpCp.type = "text";
      inpCp.id = `techBaseCp${i}`;
      inpCp.inputMode = "numeric";
      inpCp.placeholder = "1234-567";
      inpCp.maxLength = 8;
      inpCp.autocomplete = "postal-code";
      cpLab.appendChild(sp2);
      cpLab.appendChild(inpCp);
      inpCp.classList.add("postal-input");
      inpCp.addEventListener("input", () => {
        const digits = inpCp.value.replace(/\D/g, "").slice(0, 7);
        if (digits.length > 4) inpCp.value = `${digits.slice(0, 4)}-${digits.slice(4)}`;
        else inpCp.value = digits;
        applyPostalZoneClassesToInput(inpCp);
      });

      wrap.appendChild(nameLab);
      wrap.appendChild(cpLab);
      techNamesFields.appendChild(wrap);
    }
  }

  function openConfigModal() {
    ensureTechNameFields();
    const { names, baseCps, googleMapsApiKey: gkey } = loadConfig();
    for (let i = 0; i < TECH_COUNT; i++) {
      const el = document.getElementById(`techName${i}`);
      if (el) el.value = names[i] ?? "";
      const cpEl = document.getElementById(`techBaseCp${i}`);
      if (cpEl) cpEl.value = (baseCps[i] || "").trim();
    }
    const gEl = document.getElementById("fieldGoogleMapsApiKey");
    if (gEl) gEl.value = gkey || "";
    for (let i = 0; i < TECH_COUNT; i++) {
      const cpEl = document.getElementById(`techBaseCp${i}`);
      if (cpEl) applyPostalZoneClassesToInput(cpEl);
    }
    configModal.showModal();
    const first = document.getElementById("techName0");
    if (first) first.focus();
  }

  btnConfigTech?.addEventListener("click", openConfigModal);
  btnConfigClose?.addEventListener("click", () => configModal?.close());

  configForm?.addEventListener("submit", (e) => {
    e.preventDefault();
    const names = [];
    const baseCps = [];
    for (let i = 0; i < TECH_COUNT; i++) {
      const el = document.getElementById(`techName${i}`);
      names.push(el && el.value ? el.value.trim() : "");
      const cpEl = document.getElementById(`techBaseCp${i}`);
      let raw = cpEl && cpEl.value ? cpEl.value.replace(/\D/g, "").slice(0, 7) : "";
      const cpFmt = raw.length > 4 ? `${raw.slice(0, 4)}-${raw.slice(4)}` : raw.length > 0 ? raw : "";
      baseCps.push(isValidPostalPt(cpFmt) ? cpFmt : "");
    }
    const gEl = document.getElementById("fieldGoogleMapsApiKey");
    const googleMapsApiKey = gEl && gEl.value ? gEl.value.trim() : "";
    postalPairDistanceCache.clear();
    saveConfig({ names, baseCps, googleMapsApiKey });
    configModal.close();
    renderGrid({});
  });

  fieldPostal?.addEventListener("input", () => {
    const digits = fieldPostal.value.replace(/\D/g, "").slice(0, 7);
    if (digits.length > 4) fieldPostal.value = `${digits.slice(0, 4)}-${digits.slice(4)}`;
    else fieldPostal.value = digits;
    applyPostalZoneClassesToInput(fieldPostal);
  });

  if (fieldWizardPostal) {
    fieldWizardPostal.addEventListener("input", () => {
      const digits = fieldWizardPostal.value.replace(/\D/g, "").slice(0, 7);
      if (digits.length > 4) fieldWizardPostal.value = `${digits.slice(0, 4)}-${digits.slice(4)}`;
      else fieldWizardPostal.value = digits;
      applyPostalZoneClassesToInput(fieldWizardPostal);
    });
  }

  wireSlotSpanControls();

  async function showMainAppAfterAuth() {
    const authErr = document.getElementById("authError");
    setAuthFeedback(authErr, "hidden", "");

    const authScreen = document.getElementById("authScreen");
    const mainApp = document.getElementById("mainApp");
    if (authScreen) authScreen.hidden = true;
    if (mainApp) mainApp.hidden = false;
    monthInput.value = thisMonthValue();

    const syncBanner = document.getElementById("agendaSyncBanner");
    if (syncBanner) {
      syncBanner.hidden = false;
      syncBanner.classList.remove("sync-banner--error");
      syncBanner.textContent = "A sincronizar com a nuvem…";
    }

    try {
      await promiseWithTimeout(
        (async () => {
          await hydrateConfigFromCloud();
          await hydrateBookingsFromCloud();
          await maybeMigrateLocalStorageToCloud();
        })(),
        60000,
        "Timeout ao ler dados. Verifica a rede, RLS (agenda_bookings, agenda_user_settings) e permissões da app «agenda»."
      );
    } catch (e) {
      console.error(e);
      if (syncBanner) {
        syncBanner.hidden = false;
        syncBanner.classList.add("sync-banner--error");
        syncBanner.textContent = e.message || "Erro ao carregar dados da nuvem.";
      }
      renderGrid({ jumpToTodayWeek: true });
      return;
    }
    try {
      await subscribeRealtimeAgenda();
    } catch (e) {
      console.error(e);
    }
    if (syncBanner) {
      syncBanner.hidden = true;
      syncBanner.classList.remove("sync-banner--error");
    }
    renderGrid({ jumpToTodayWeek: true });
  }

  async function startAgenda() {
    const cfg = window.AGENDA_CONFIG;
    if (isLocalOnly()) {
      supabaseClient = null;
      loadBookingsFromLocal();
      loadConfigFromLocalStorageIntoSnapshot();
      const authScreen = document.getElementById("authScreen");
      const mainApp = document.getElementById("mainApp");
      if (authScreen) authScreen.hidden = true;
      if (mainApp) mainApp.hidden = false;
      const banner = document.getElementById("localModeBanner");
      if (banner) banner.hidden = false;
      const btnLogout = document.getElementById("btnLogout");
      if (btnLogout) btnLogout.hidden = true;
      monthInput.value = thisMonthValue();
      renderGrid({ jumpToTodayWeek: true });
      return;
    }

    const key = cfg?.supabaseKey ? String(cfg.supabaseKey).trim() : "";
    const invalidConfig = !cfg?.supabaseUrl || !key || key.includes("COLOCA_AQUI");
    const configHint = !cfg
      ? "O config não carregou (404). No deploy unificado, site/agenda/index.html deve ter <script src=\"../config.js\"></script> e existir site/config.js na raiz. Corre ./scripts/sync-site.sh e publica de novo a pasta site/."
      : "Configura agenda-servicos/config.js com supabaseUrl e supabaseKey (igual ao portal).";
    if (invalidConfig) {
      setAuthFeedback(document.getElementById("authError"), "error", configHint);
    }

    const url = invalidConfig ? "" : String(cfg.supabaseUrl).trim().replace(/\/+$/, "");

    /**
     * Inicialização em paralelo: o listener do formulário tem de existir já (com preventDefault).
     * Antes, o await import() podia falhar ou demorar e o clique em «Entrar» não fazia nada.
     */
    const initSupabaseDone = invalidConfig
      ? Promise.resolve()
      : (async () => {
      const { createClient } = await import("https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2/+esm");
      const persist = cfg.persistSession !== false;
      supabaseClient = createClient(url, key, {
        auth: { persistSession: persist, autoRefreshToken: persist },
        realtime: { params: { eventsPerSecond: 30 } },
      });

      supabaseClient.auth.onAuthStateChange((event, session) => {
        if (session?.access_token) {
          setRealtimeAuthFromSession(session);
        }
        if (event === "TOKEN_REFRESHED" && session?.access_token) {
          void subscribeRealtimeAgenda();
        }
        if (event === "SIGNED_OUT") {
          unsubscribeRealtimeAgenda();
          __bookingStore = { __schema: 2 };
          __configSnapshot = defaultAgendaConfig();
          const authScreen = document.getElementById("authScreen");
          const mainApp = document.getElementById("mainApp");
          if (mainApp) mainApp.hidden = true;
          if (authScreen) authScreen.hidden = false;
        }
      });

      const {
        data: { session },
      } = await supabaseClient.auth.getSession();
      if (session) {
        await showMainAppAfterAuth();
      } else {
        const authScreen = document.getElementById("authScreen");
        const mainApp = document.getElementById("mainApp");
        if (authScreen) authScreen.hidden = false;
        if (mainApp) mainApp.hidden = true;
      }
        })();

    async function runAgendaLogin() {
      const errEl = document.getElementById("authError");
      if (invalidConfig) {
        setAuthFeedback(errEl, "error", configHint);
        return;
      }
      setAuthFeedback(errEl, "loading", "A ligar ao servidor…");
      const btnEntrar = document.getElementById("btnEntrar");
      if (btnEntrar) {
        btnEntrar.disabled = true;
        btnEntrar.setAttribute("aria-busy", "true");
      }
      try {
        try {
          await promiseWithTimeout(
            initSupabaseDone,
            45000,
            "Timeout ao inicializar o Supabase. Verifica rede, bloqueios do browser e HTTP referrers da chave API (Supabase → API Keys)."
          );
        } catch (err) {
          console.error(err);
          setAuthFeedback(
            errEl,
            "error",
            err?.message ||
              "Não foi possível ligar ao Supabase. Verifica a rede, o config.js e os HTTP referrers da chave no dashboard."
          );
          return;
        }
        if (!supabaseClient) {
          setAuthFeedback(errEl, "error", "Cliente Supabase não inicializado.");
          return;
        }
        const email = document.getElementById("authEmail")?.value?.trim();
        const password = document.getElementById("authPassword")?.value;
        setAuthFeedback(errEl, "loading", "A iniciar sessão…");
        let signResult;
        try {
          signResult = await promiseWithTimeout(
            supabaseClient.auth.signInWithPassword({ email, password }),
            45000,
            "Timeout ao iniciar sessão. Verifica rede e se o domínio Netlify está nos referrers da chave API."
          );
        } catch (timeoutErr) {
          setAuthFeedback(
            errEl,
            "error",
            timeoutErr?.message || "Timeout ao iniciar sessão."
          );
          return;
        }
        const { error: signErr } = signResult;
        if (signErr) {
          setAuthFeedback(errEl, "error", signErr.message);
          return;
        }
        await showMainAppAfterAuth();
      } catch (err) {
        console.error(err);
        setAuthFeedback(errEl, "error", err?.message || "Erro ao iniciar sessão.");
      } finally {
        if (btnEntrar) {
          btnEntrar.disabled = false;
          btnEntrar.removeAttribute("aria-busy");
        }
      }
    }

    document.getElementById("authForm")?.addEventListener("submit", (e) => {
      e.preventDefault();
    });
    document.getElementById("btnEntrar")?.addEventListener("click", (e) => {
      e.preventDefault();
      void runAgendaLogin();
    });
    for (const id of ["authEmail", "authPassword"]) {
      document.getElementById(id)?.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
          e.preventDefault();
          void runAgendaLogin();
        }
      });
    }

    document.getElementById("btnLogout")?.addEventListener("click", async () => {
      try {
        await initSupabaseDone;
        if (supabaseClient) await supabaseClient.auth.signOut();
      } catch (err) {
        console.error(err);
      }
    });

    try {
      await initSupabaseDone;
    } catch (err) {
      console.error(err);
      setAuthFeedback(
        document.getElementById("authError"),
        "error",
        err?.message ||
          "Erro ao inicializar o Supabase. Confirma config.js, chave API e HTTP referrers para este domínio Netlify."
      );
    }
  }
})();
