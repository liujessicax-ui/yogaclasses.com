/**
 * schedule-data.js — shared schedule loader for the public site.
 *
 * Single source of truth is the Google Sheet, served by the Apps Script at
 * ?action=schedule. This module fetches it, caches the last-good copy in
 * localStorage, and falls back to a baked-in constant so the pages still
 * render if the script is cold or unreachable.
 *
 * Exposes window.YogaSchedule with:
 *   load(url)            -> Promise<{classes, exceptions, source}>  (fetch + cache)
 *   current()           -> {classes, exceptions}  (sync: cache or fallback)
 *   upcomingOccurrences(cls, exceptions, windowDays) -> [{date, startHour, startMin, durationMins}]
 *   formatDate(date), formatTimeRange(h, m, durationMins)
 *
 * All times are Pacific wall-clock; pacificDate() handles PST/PDT (DST) the
 * same way the original signup.html logic did — do not switch to UTC.
 */
(function (global) {
  'use strict';

  var LA_TZ = 'America/Los_Angeles';
  var CACHE_KEY = 'yogaScheduleCache';

  // ---- FALLBACK ----------------------------------------------------------
  // Used ONLY when the Apps Script endpoint is unreachable AND there's no
  // cached copy. The Google Sheet is the real source of truth; keep this in
  // rough sync but don't rely on it for day-to-day changes.
  var FALLBACK = {
    classes: [
      { id: 'sunday-online', label: 'Sunday Evening — Online via Zoom', day: 0, startTime: '18:00', durationMins: 75, type: 'online', location: '', capacity: null, tags: ['Open to Everyone'], props: ['Yoga mat', 'Strap', 'Two blocks', 'Wall space', 'Yoga chair (ideal)', 'Bolster (ideal)'], active: true },
      { id: 'tuesday-ccv', label: 'Tuesday Evening — CCV Clubhouse (In Person)', day: 2, startTime: '18:00', durationMins: 75, type: 'inperson', location: 'CCV Clubhouse', capacity: 10, tags: ['CCV Residents Only', 'In Person'], props: ['Yoga mat', 'Two blocks', 'Strap'], active: true },
      { id: 'wednesday-restorative', label: 'Wednesday Evening — Restorative Yoga (Online)', day: 3, startTime: '20:00', durationMins: 75, type: 'online', location: '', capacity: null, tags: ['Restorative', 'Open to Everyone'], props: ['Yoga mat', 'Bolster', 'Two blocks', 'Two blankets', 'Strap', 'Wall space', 'Yoga chair (ideal)'], active: true }
    ],
    exceptions: []
  };

  // ---- small helpers -----------------------------------------------------
  function pad2(n) { n = parseInt(n, 10) || 0; return (n < 10 ? '0' : '') + n; }

  function parseHM(v) {
    var m = String(v == null ? '' : v).match(/(\d{1,2}):(\d{2})/);
    if (m) return { h: parseInt(m[1], 10), m: parseInt(m[2], 10) };
    var n = parseInt(v, 10);
    if (!isNaN(n) && n >= 0 && n <= 23) return { h: n, m: 0 };
    return { h: 18, m: 0 };
  }

  function asArray(v) {
    if (Array.isArray(v)) return v.map(function (x) { return String(x).trim(); }).filter(Boolean);
    return String(v == null ? '' : v).split(',').map(function (x) { return x.trim(); }).filter(Boolean);
  }

  function isoFromParts(y, mo, da) { return y + '-' + pad2(mo) + '-' + pad2(da); }

  function dateFromIso(iso) {
    var m = String(iso || '').match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (!m) return null;
    return { year: +m[1], month: +m[2], day: +m[3] };
  }

  // Normalize a value to a "YYYY-MM-DD" string, or '' if it isn't a date.
  function isoOrEmpty(v) {
    var m = String(v == null ? '' : v).match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    return m ? (m[1] + '-' + pad2(+m[2]) + '-' + pad2(+m[3])) : '';
  }

  function findEx(exceptions, classId, iso) {
    for (var i = 0; i < exceptions.length; i++) {
      if (exceptions[i].classId === classId && exceptions[i].date === iso) return exceptions[i];
    }
    return null;
  }

  // ---- Pacific time (DST-aware) — ported from signup.html ----------------
  function getNowInPacific() {
    var now = new Date();
    var parts = {};
    new Intl.DateTimeFormat('en-US', {
      timeZone: LA_TZ,
      year: 'numeric', month: 'numeric', day: 'numeric',
      hour: 'numeric', minute: 'numeric', second: 'numeric',
      hour12: false, weekday: 'short'
    }).formatToParts(now).forEach(function (p) { parts[p.type] = p.value; });
    return {
      year: +parts.year,
      month: +parts.month,
      day: +parts.day,
      hour: +parts.hour === 24 ? 0 : +parts.hour,
      minute: +parts.minute,
      _date: now
    };
  }

  // Create a Date for a specific Pacific date+time (DST-aware).
  function pacificDate(year, month, day, hour, minute) {
    var base = year + '-' + pad2(month) + '-' + pad2(day) + 'T' + pad2(hour) + ':' + pad2(minute) + ':00';
    for (var i = 0; i < 2; i++) {
      var offset = i === 0 ? '-08:00' : '-07:00';
      var candidate = new Date(base + offset);
      var checkHour = +new Intl.DateTimeFormat('en-US', {
        timeZone: LA_TZ, hour: 'numeric', hour12: false
      }).formatToParts(candidate).find(function (p) { return p.type === 'hour'; }).value;
      if (checkHour === hour || (checkHour === 24 && hour === 0)) return candidate;
    }
    return new Date(base + '-08:00');
  }

  function formatDate(date) {
    return date.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric', timeZone: LA_TZ });
  }

  function fmtTime(h, m) {
    var ampm = h >= 12 ? 'PM' : 'AM';
    var hh = h % 12; if (hh === 0) hh = 12;
    return hh + ':' + pad2(m) + ' ' + ampm;
  }

  function formatTimeRange(startHour, startMin, durationMins) {
    var endTotal = startHour * 60 + startMin + (parseInt(durationMins, 10) || 0);
    var eh = Math.floor(endTotal / 60) % 24, em = endTotal % 60;
    return fmtTime(startHour, startMin) + ' – ' + fmtTime(eh, em) + ' PT';
  }

  // ---- normalize / cache -------------------------------------------------
  function normalizeClass(c) {
    var t = parseHM(c.startTime != null ? c.startTime : (pad2(c.startH) + ':' + pad2(c.startM)));
    var cap = (c.capacity === '' || c.capacity == null) ? null : (parseInt(c.capacity, 10));
    return {
      id: String(c.id || '').trim(),
      label: String(c.label || '').trim(),
      day: parseInt(c.day, 10) || 0,
      startHour: t.h,
      startMin: t.m,
      durationMins: parseInt(c.durationMins, 10) || 75,
      type: String(c.type || 'online').trim().toLowerCase(),
      location: String(c.location || '').trim(),
      capacity: isNaN(cap) ? null : cap,
      tags: asArray(c.tags),
      props: asArray(c.props),
      active: c.active !== false && String(c.active).toLowerCase() !== 'false',
      // Blank = recurring weekly on `day`; a date = a single, non-recurring occurrence.
      oneOffDate: isoOrEmpty(c.oneOffDate)
    };
  }

  function normalize(data) {
    var classes = (data && Array.isArray(data.classes) ? data.classes : []).map(normalizeClass).filter(function (c) { return c.id; });
    var exceptions = (data && Array.isArray(data.exceptions) ? data.exceptions : []).map(function (e) {
      return {
        classId: String(e.classId || '').trim(),
        date: String(e.date || '').trim(),
        status: String(e.status || '').trim().toLowerCase(),
        newDate: String(e.newDate || '').trim(),
        newStartTime: String(e.newStartTime || '').trim(),
        note: String(e.note || '')
      };
    }).filter(function (e) { return e.classId && e.date; });
    return { classes: classes, exceptions: exceptions };
  }

  function readCache() {
    try { return JSON.parse(global.localStorage.getItem(CACHE_KEY)); } catch (e) { return null; }
  }
  function writeCache(data) {
    try { global.localStorage.setItem(CACHE_KEY, JSON.stringify(data)); } catch (e) {}
  }

  function current() {
    var c = readCache();
    if (c && Array.isArray(c.classes) && c.classes.length) return normalize(c);
    return normalize(FALLBACK);
  }

  function withSource(src, data) {
    return { classes: data.classes, exceptions: data.exceptions, source: src };
  }

  function load(url) {
    return new Promise(function (resolve) {
      if (!url) { resolve(withSource('fallback', current())); return; }
      var full = url + (url.indexOf('?') >= 0 ? '&' : '?') + 'action=schedule';
      global.fetch(full)
        .then(function (r) { return r.json(); })
        .then(function (j) {
          if (j && j.status === 'ok' && Array.isArray(j.classes) && j.classes.length) {
            var data = { classes: j.classes, exceptions: j.exceptions || [] };
            writeCache(data);
            resolve(withSource('network', normalize(data)));
          } else {
            resolve(withSource('cache-or-fallback', current()));
          }
        })
        .catch(function () { resolve(withSource('cache-or-fallback', current())); });
    });
  }

  // ---- occurrence expansion ---------------------------------------------
  // Return upcoming dated occurrences of a class within windowDays, applying
  // exceptions: cancelled/moved removes the original date; moved/extra add a
  // new date (with an optional new start time).
  function upcomingOccurrences(cls, exceptions, windowDays) {
    if (!cls.active) return [];
    exceptions = exceptions || [];
    var nowUTC = new Date();
    var pac = getNowInPacific();
    var windowEnd = new Date(nowUTC.getTime() + windowDays * 86400000);
    var results = [];
    var seen = {};

    // ignoreWindow=true lets one-off classes be signed up for any time before the
    // event (not just within windowDays) — they're announced ahead of time, unlike
    // the rolling weekly series which only open ~windowDays out.
    function consider(y, mo, da, h, m, ignoreWindow) {
      var classStart = pacificDate(y, mo, da, h, m);
      var signupCutoff = new Date(classStart.getTime() + 15 * 60000);
      if (signupCutoff > nowUTC && (ignoreWindow || classStart <= windowEnd)) {
        var key = classStart.getTime();
        if (seen[key]) return;
        seen[key] = true;
        results.push({ date: classStart, startHour: h, startMin: m, durationMins: cls.durationMins });
      }
    }

    if (cls.oneOffDate) {
      // One-off: available to sign up from now right up to its date (no weekly
      // scan, no windowDays cap), unless that date is cancelled or moved away.
      var od = dateFromIso(cls.oneOffDate);
      if (od) {
        var oex = findEx(exceptions, cls.id, cls.oneOffDate);
        if (!(oex && (oex.status === 'cancelled' || oex.status === 'moved'))) {
          consider(od.year, od.month, od.day, cls.startHour, cls.startMin, true);
        }
      }
    } else {
      for (var offset = 0; offset <= windowDays + 1; offset++) {
        var scan = new Date(pac.year, pac.month - 1, pac.day + offset);
        if (scan.getDay() !== cls.day) continue;
        var iso = isoFromParts(scan.getFullYear(), scan.getMonth() + 1, scan.getDate());
        var ex = findEx(exceptions, cls.id, iso);
        if (ex && (ex.status === 'cancelled' || ex.status === 'moved')) continue;
        consider(scan.getFullYear(), scan.getMonth() + 1, scan.getDate(), cls.startHour, cls.startMin);
      }
    }

    exceptions.forEach(function (ex) {
      if (ex.classId !== cls.id) return;
      var targetIso = '', t = { h: cls.startHour, m: cls.startMin };
      if (ex.status === 'moved' && ex.newDate) { targetIso = ex.newDate; if (ex.newStartTime) t = parseHM(ex.newStartTime); }
      else if (ex.status === 'extra') { targetIso = ex.date; if (ex.newStartTime) t = parseHM(ex.newStartTime); }
      if (!targetIso) return;
      var d = dateFromIso(targetIso);
      if (!d) return;
      // A moved/extra date for a one-off keeps its "signable anytime" behavior.
      consider(d.year, d.month, d.day, t.h, t.m, !!cls.oneOffDate);
    });

    results.sort(function (a, b) { return a.date - b.date; });
    return results;
  }

  global.YogaSchedule = {
    LA_TZ: LA_TZ,
    FALLBACK: FALLBACK,
    load: load,
    current: current,
    normalize: normalize,
    upcomingOccurrences: upcomingOccurrences,
    formatDate: formatDate,
    formatTimeRange: formatTimeRange,
    pacificDate: pacificDate,
    getNowInPacific: getNowInPacific
  };
})(window);
