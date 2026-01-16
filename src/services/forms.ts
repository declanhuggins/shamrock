// Form builders to scaffold required questions. Idempotent: only seeds when the form is empty.

namespace FormService {
  function clearItems(form: GoogleAppsScript.Forms.Form) {
    const items = form.getItems();

    // First, strip navigation from choice items so page break dependencies don't block deletion.
    items.forEach((item) => {
      const type = item.getType && item.getType();
      if (type === FormApp.ItemType.MULTIPLE_CHOICE || type === FormApp.ItemType.LIST) {
        try {
          const mc = (type === FormApp.ItemType.MULTIPLE_CHOICE ? item.asMultipleChoiceItem?.() : item.asListItem?.()) as any;
          if (!mc) return;
          const choices = mc.getChoices().map((c: GoogleAppsScript.Forms.Choice) => mc.createChoice(c.getValue()));
          mc.setChoices(choices);
        } catch (err) {
          Log.warn(`Unable to strip navigation from item '${item.getTitle?.() || ''}': ${err}`);
        }
      }
    });

    for (let i = items.length - 1; i >= 0; i--) {
      try {
        form.deleteItem(items[i]);
      } catch (err) {
        Log.warn(`Unable to delete form item '${items[i].getTitle?.() || ''}': ${err}`);
      }
    }
  }

  function removeItemsByTitle(form: GoogleAppsScript.Forms.Form, titles: string[]) {
    const titleSet = new Set(titles.map((t) => t.toLowerCase()));
    form.getItems().forEach((item) => {
      const t = (item.getTitle && item.getTitle()) || '';
      if (titleSet.has(String(t).toLowerCase())) {
        try {
          form.deleteItem(item);
          Log.info(`Removed redundant form item titled '${t}'`);
        } catch (err) {
          Log.warn(`Unable to remove item titled '${t}'. Error: ${err}`);
        }
      }
    });
  }

  function seedIfEmpty(form: GoogleAppsScript.Forms.Form, builder: (f: GoogleAppsScript.Forms.Form) => void, label: string) {
    if (form.getItems().length === 0) {
      Log.info(`Seeding form items for ${label}`);
      builder(form);
    } else {
      Log.info(`Form ${label} already has items; not modifying questions.`);
    }
  }

  interface CadetOption {
    label: string;
    asYear: string;
    flight: string;
    university: string;
  }

  interface CadetGroups {
    byFlight: Record<string, Record<string, string[]>>; // flight -> AS -> labels (excludes crosstown for Mando)
    byFlightAll: Record<string, Record<string, string[]>>; // flight -> AS -> labels (includes crosstown for LLAB)
    byCrosstown: Record<string, Record<string, string[]>>; // university -> AS -> labels
    allByAs: Record<string, string[]>; // AS -> labels
    nonAbroadByAs: Record<string, string[]>; // AS -> labels (exclude flight Abroad)
  }

  function normalizeToList(value: string, options: string[]): string {
    const v = String(value || '').trim();
    if (!v) return '';
    const lc = v.toLowerCase();
    const exact = options.find((o) => o.toLowerCase() === lc);
    if (exact) return exact;
    const prefix = options.find((o) => lc.startsWith(o.toLowerCase()));
    if (prefix) return prefix;
    const contains = options.find((o) => lc.includes(o.toLowerCase()));
    if (contains) return contains;
    return '';
  }

  function buildCadetGroups(): CadetGroups {
    const groups: CadetGroups = { byFlight: {}, byFlightAll: {}, byCrosstown: {}, allByAs: {}, nonAbroadByAs: {} };
    try {
      const backendId = Config.getBackendId();
      const sheet = SheetUtils.getSheet(backendId, 'Directory Backend');
      if (!sheet) return groups;
      const table = SheetUtils.readTable(sheet);
      table.rows.forEach((r) => {
        const as = String(r['as_year'] || '').trim() || 'Unknown';
        const flightRaw = String(r['flight'] || '').trim();
        const normalizedFlight = normalizeToList(flightRaw, Arrays.FLIGHTS);
        const flight = normalizedFlight || flightRaw;
        const university = String(r['university'] || '').trim();
        const label = `${r['last_name'] || ''}, ${r['first_name'] || ''}`.trim();

        // all cadets grouped by AS
        groups.allByAs[as] = groups.allByAs[as] || [];
        groups.allByAs[as].push(label);

        const isAbroad = normalizedFlight === 'Abroad' || flightRaw.toLowerCase() === 'abroad';
        if (!isAbroad) {
          groups.nonAbroadByAs[as] = groups.nonAbroadByAs[as] || [];
          groups.nonAbroadByAs[as].push(label);
        }

        const uniLc = university.toLowerCase().trim();
        // Treat Trine and Valpo/Valparaiso as crosstown regardless of extra words
          const isCrosstown = !isAbroad && (/trine/.test(uniLc) || /valpo|valpar/.test(uniLc));

        if (flight) {
          // Full flight grouping (including crosstown) for LLAB
          groups.byFlightAll[flight] = groups.byFlightAll[flight] || {};
          groups.byFlightAll[flight][as] = groups.byFlightAll[flight][as] || [];
          groups.byFlightAll[flight][as].push(label);

          groups.byFlight[flight] = groups.byFlight[flight] || {};
          groups.byFlight[flight][as] = groups.byFlight[flight][as] || [];
          // Exclude crosstown cadets from byFlight (they go to byCrosstown for Mando)
          if (!isAbroad && !isCrosstown) {
            groups.byFlight[flight][as].push(label);
          }
        }

        // Add crosstown cadets to byCrosstown for Mando branch
        if (isCrosstown) {
          groups.byCrosstown[university] = groups.byCrosstown[university] || {};
          groups.byCrosstown[university][as] = groups.byCrosstown[university][as] || [];
          groups.byCrosstown[university][as].push(label);
        }
      });

      const sortValues = (m: Record<string, string[]>) => {
        Object.keys(m).forEach((k) => m[k].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' })));
      };
      sortValues(groups.allByAs);
      sortValues(groups.nonAbroadByAs);
      Object.values(groups.byFlight).forEach(sortValues);
      Object.values(groups.byFlightAll).forEach(sortValues);
      Object.values(groups.byCrosstown).forEach(sortValues);
    } catch (err) {
      Log.warn(`Unable to build cadet groups: ${err}`);
    }
    return groups;
  }

  function enforceExcusalsItemOrder(form: GoogleAppsScript.Forms.Form) {
    const desired = ['Last Name', 'First Name', 'Event', 'Reason'];
    const items = form.getItems();
    const findByTitle = (title: string) =>
      items.find((item) => String(item.getTitle?.() || '').trim().toLowerCase() === title.toLowerCase());

    desired.forEach((title, idx) => {
      const item = findByTitle(title);
      if (item) {
        try {
          form.moveItem(item, idx);
        } catch (err) {
          Log.warn(`Unable to move item ${title}: ${err}`);
        }
      }
    });
  }

  function applyDirectoryRegexValidations(form: GoogleAppsScript.Forms.Form) {
    const findTextItem = (title: string): GoogleAppsScript.Forms.TextItem | null => {
      const item = form.getItems(FormApp.ItemType.TEXT).find((i) => String(i.getTitle() || '').trim() === title);
      return item ? item.asTextItem() : null;
    };

    const classYearItem = findTextItem('Class Year (YYYY)');
    if (classYearItem) {
      try {
        const classYearPattern = /^\d{4}$/;
        const validation = FormApp.createTextValidation().setHelpText('Enter a 4-digit year (YYYY)').requireTextMatchesPattern(classYearPattern.source).build();
        classYearItem.setValidation(validation);
      } catch (err) {
        Log.warn(`Unable to apply class year validation on Directory form: ${err}`);
      }
    }

    const cipCodeItem = findTextItem('CIP Code (XX.XXXX)');
    if (cipCodeItem) {
      try {
        const cipPattern = /^\d{2}\.\d{4}$/;
        const validation = FormApp.createTextValidation().setHelpText('Format: 12.3456').requireTextMatchesPattern(cipPattern.source).build();
        cipCodeItem.setValidation(validation);
      } catch (err) {
        Log.warn(`Unable to apply CIP code validation on Directory form: ${err}`);
      }
    }

    const phoneTitles = ['Phone (+5 (555) 555-5555)', 'Phone (+1 (555) 555-5555)'];
    const phoneItem = phoneTitles.reduce<GoogleAppsScript.Forms.TextItem | null>((found, title) => found || findTextItem(title), null);
    if (phoneItem) {
      try {
        const phonePattern = /^\+\d \(\d{3}\) \d{3}-\d{4}$/;
        const validation = FormApp.createTextValidation()
          .setHelpText('Format: +5 (555) 555-5555')
          .requireTextMatchesPattern(phonePattern.source)
          .build();
        phoneItem.setValidation(validation);
      } catch (err) {
        Log.warn(`Unable to apply phone validation on Directory form: ${err}`);
      }
    }

    const photoItem = findTextItem('Photo Link (URL)');
    if (photoItem) {
      try {
        const validation = FormApp.createTextValidation().setHelpText('Enter a valid URL').requireTextIsUrl().build();
        photoItem.setValidation(validation);
      } catch (err) {
        Log.warn(`Unable to apply photo URL validation on Directory form: ${err}`);
      }
    }
  }

  export function ensureDirectoryForm(form: GoogleAppsScript.Forms.Form) {
    // Prune redundant email item; form already collects verified email via settings.
    removeItemsByTitle(form, ['University Email']);

    seedIfEmpty(form, (f) => {
      f.addTextItem().setTitle('Last Name').setRequired(true);
      f.addTextItem().setTitle('First Name').setRequired(true);
      f.addListItem().setTitle('AS Year').setChoiceValues(Arrays.AS_YEARS).setRequired(true);
      f.addTextItem().setTitle('Class Year (YYYY)').setRequired(true);
      f.addListItem().setTitle('Flight').setChoiceValues(Arrays.FLIGHTS);
      f.addListItem().setTitle('Squadron').setChoiceValues(Arrays.SQUADRONS);
      f.addListItem().setTitle('University').setChoiceValues(Arrays.UNIVERSITIES).setRequired(true);
      f.addTextItem().setTitle('Phone (+5 (555) 555-5555)').setRequired(true);
      f.addListItem().setTitle('Dorm').setChoiceValues(Arrays.DORMS);
      f.addTextItem().setTitle('Home Town').setRequired(true);
      f.addListItem().setTitle('Home State').setChoiceValues(Arrays.HOME_STATES).setRequired(true);
      f.addDateItem().setTitle('DOB (MM/DD/YYYY)').setRequired(true);
      f.addListItem().setTitle('CIP Broad Area').setChoiceValues(Arrays.CIP_BROAD_AREAS);
      f.addTextItem().setTitle('CIP Code (XX.XXXX)');
      f.addListItem().setTitle('Desired/Assigned AFSC').setChoiceValues(Arrays.AFSC_OPTIONS);
      f.addListItem().setTitle('Flight Path Status').setChoiceValues(Arrays.FLIGHT_PATH_STATUSES);
      f.addTextItem().setTitle('Photo Link (URL)');
      f.addParagraphTextItem().setTitle('Notes');
    }, 'Directory Form');

    applyDirectoryRegexValidations(form);
  }

  export function ensureAttendanceForm(form: GoogleAppsScript.Forms.Form) {
    // If empty, build the full form; otherwise just keep the Event list up to date.
    if (form.getItems().length === 0) {
      rebuildAttendanceForm(form);
      return;
    }

    refreshAttendanceFormEventChoices(form);
  }

  export function rebuildAttendanceForm(form: GoogleAppsScript.Forms.Form) {
    // Rebuild to enforced multi-page structure with branching by event type.
    Log.info('Attendance form: start rebuild');
    clearItems(form);
    Log.info('Attendance form: cleared existing items');

    const cadets = buildCadetGroups();
    Log.info(`Attendance form: cadet groups built flights=${Object.keys(cadets.byFlight).length} crosstown=${Object.keys(cadets.byCrosstown).length}`);

    // Section 1: respondent info + event
    form.addTextItem().setTitle('Name').setRequired(true);
    form.addListItem().setTitle('Event').setRequired(true);

    // Section 2: Mando branch selector
    form.addPageBreakItem().setTitle('Mando Branch');
    const mandoFlightItem = form.addMultipleChoiceItem().setTitle('Flight / Crosstown (Mando)').setRequired(true);
    const mandoFlights = [...Arrays.FLIGHTS.filter((f) => f !== 'Abroad'), 'Trine', 'Valparaiso'];
    const mandoFlightPages: Record<string, GoogleAppsScript.Forms.PageBreakItem> = {};
    mandoFlights.forEach((fName) => {
      const page = form
        .addPageBreakItem()
        .setTitle(`Cadets for ${fName} (Mando)`)
        .setGoToPage(FormApp.PageNavigationType.SUBMIT);
      mandoFlightPages[fName] = page;

      let groupMap: Record<string, string[]> = {};
      if (fName === 'Trine' || fName === 'Valparaiso') {
        const matchKey = Object.keys(cadets.byCrosstown).find((k) => k.toLowerCase().includes(fName.toLowerCase())) || '';
        groupMap = (matchKey && cadets.byCrosstown[matchKey]) || {};
      } else {
        groupMap = cadets.byFlight[fName] || {};
      }

      Object.keys(groupMap)
        .sort((a, b) => b.localeCompare(a, undefined, { sensitivity: 'base' }))
        .forEach((as) => {
          const opts = groupMap[as];
          if (!opts || !opts.length) return;
          form.addCheckboxItem().setTitle(`Cadets (${fName}) AS ${as} (Mando)`).setChoiceValues(opts);
        });
    });
    mandoFlightItem.setChoices(mandoFlights.map((f) => mandoFlightItem.createChoice(f, mandoFlightPages[f])));
    Log.info(`Attendance form: Mando flight pages=${mandoFlights.length}`);

    // Section 11: LLAB branch selector
    form.addPageBreakItem().setTitle('LLAB Branch');
    const llabFlightItem = form.addMultipleChoiceItem().setTitle('Flight (LLAB)').setRequired(true);
    const llabFlights = Arrays.FLIGHTS.filter((f) => f !== 'Abroad');
    const llabFlightPages: Record<string, GoogleAppsScript.Forms.PageBreakItem> = {};
    llabFlights.forEach((fName) => {
      const page = form
        .addPageBreakItem()
        .setTitle(`Cadets for ${fName} (LLAB)`)
        .setGoToPage(FormApp.PageNavigationType.SUBMIT);
      llabFlightPages[fName] = page;

      const groupMap = cadets.byFlightAll[fName] || cadets.byFlight[fName] || {};
      Object.keys(groupMap)
        .sort((a, b) => b.localeCompare(a, undefined, { sensitivity: 'base' }))
        .forEach((as) => {
          const opts = groupMap[as];
          if (!opts || !opts.length) return;
          form.addCheckboxItem().setTitle(`Cadets (${fName}) AS ${as} (LLAB)`).setChoiceValues(opts);
        });
    });
    llabFlightItem.setChoices(llabFlights.map((f) => llabFlightItem.createChoice(f, llabFlightPages[f])));
    Log.info(`Attendance form: LLAB flight pages=${llabFlights.length}`);

    // Section 18: Secondary branch (non-abroad) -> submit
    const secondaryStart = form
      .addPageBreakItem()
      .setTitle('Secondary Branch')
      .setGoToPage(FormApp.PageNavigationType.SUBMIT);
    Object.keys(cadets.nonAbroadByAs)
      .sort((a, b) => b.localeCompare(a, undefined, { sensitivity: 'base' }))
      .forEach((as) => {
        const opts = cadets.nonAbroadByAs[as];
        if (!opts || !opts.length) return;
        form.addCheckboxItem().setTitle(`Cadets (Secondary) AS ${as}`).setChoiceValues(opts);
      });

    // Section 19: Other branch (all cadets) -> submit
    const fallbackStart = form
      .addPageBreakItem()
      .setTitle('Attendance Branch')
      .setGoToPage(FormApp.PageNavigationType.SUBMIT);
    Object.keys(cadets.allByAs)
      .sort((a, b) => b.localeCompare(a, undefined, { sensitivity: 'base' }))
      .forEach((as) => {
        const opts = cadets.allByAs[as];
        if (!opts || !opts.length) return;
        form.addCheckboxItem().setTitle(`Cadets (All) AS ${as}`).setChoiceValues(opts);
      });

    // Wire Event choices now that the branch targets exist.
    refreshAttendanceFormEventChoices(form);
    Log.info('Attendance form: completed rebuild');
  }

  export function refreshAttendanceFormEventChoices(form: GoogleAppsScript.Forms.Form) {
    let eventQuestion: GoogleAppsScript.Forms.ListItem | GoogleAppsScript.Forms.MultipleChoiceItem | null = null;

    const listItems = form.getItems(FormApp.ItemType.LIST);
    for (const item of listItems) {
      try {
        if (String(item.getTitle() || '').trim() === 'Event') {
          eventQuestion = item.asListItem();
          break;
        }
      } catch {
        // ignore
      }
    }

    if (!eventQuestion) {
      const mcItems = form.getItems(FormApp.ItemType.MULTIPLE_CHOICE);
      for (const item of mcItems) {
        try {
          if (String(item.getTitle() || '').trim() === 'Event') {
            eventQuestion = item.asMultipleChoiceItem();
            break;
          }
        } catch {
          // ignore
        }
      }
    }

    if (!eventQuestion) {
      Log.warn('Attendance form: cannot refresh events; no list/multiple-choice item titled "Event"');
      return;
    }

    const pages = form.getItems(FormApp.ItemType.PAGE_BREAK);
    const findPage = (title: string) => {
      const match = pages.find((p) => {
        try {
          return String(p.getTitle() || '').trim() === title;
        } catch {
          return false;
        }
      });
      return match ? match.asPageBreakItem() : null;
    };

    const mandoStart = findPage('Mando Branch');
    const llabStart = findPage('LLAB Branch');
    const secondaryStart = findPage('Secondary Branch');
    const fallbackStart = findPage('Attendance Branch');

    const eventChoices: GoogleAppsScript.Forms.Choice[] = [];
    try {
      const backendId = Config.getBackendId();
      const eventsSheet = backendId ? SheetUtils.getSheet(backendId, 'Events Backend') : null;
      if (eventsSheet) {
        const eventsTable = SheetUtils.readTable(eventsSheet);
        eventsTable.rows.forEach((r) => {
          const name = r['display_name'] || r['attendance_column_label'] || r['event_id'];
          if (!name) return;
          const type = String(r['event_type'] || '').toLowerCase();
          let target: GoogleAppsScript.Forms.PageBreakItem | GoogleAppsScript.Forms.PageNavigationType | null = fallbackStart;
          if (type.includes('llab')) target = llabStart;
          else if (type.includes('mando')) target = mandoStart;
          else if (type.includes('secondary')) target = secondaryStart;

          if (target) {
            eventChoices.push((eventQuestion as any).createChoice(name, target));
          } else {
            eventChoices.push((eventQuestion as any).createChoice(name, FormApp.PageNavigationType.SUBMIT));
          }
        });
      }
    } catch (err) {
      Log.warn(`Unable to populate attendance form events: ${err}`);
    }

    if (!eventChoices.length) {
      if (fallbackStart) eventChoices.push(eventQuestion.createChoice('(set up events first)', fallbackStart));
      else eventChoices.push(eventQuestion.createChoice('(set up events first)', FormApp.PageNavigationType.SUBMIT));
    }
    (eventQuestion as any).setChoices(eventChoices);
    Log.info(`Attendance form: refreshed event choices count=${eventChoices.length}`);
  }

  export function ensureExcusalsForm(form: GoogleAppsScript.Forms.Form) {
    // Prune redundant/legacy items; form already collects verified email via settings.
    removeItemsByTitle(form, ['University Email']);
    removeItemsByTitle(form, ['Event']);
    removeItemsByTitle(form, ['Other Event (if "Other")']);

    seedIfEmpty(form, (f) => {
      f.addTextItem().setTitle('Last Name').setRequired(true);
      f.addTextItem().setTitle('First Name').setRequired(true);
      f.addCheckboxItem().setTitle('Event').setRequired(true);
      f.addParagraphTextItem().setTitle('Reason').setRequired(true);
    }, 'Excusals Form');

    refreshExcusalsFormEventChoices(form);
    enforceExcusalsItemOrder(form);
  }

  export function refreshExcusalsFormEventChoices(form: GoogleAppsScript.Forms.Form) {
    // Find or create the Event question as checkboxes (allow multiple selections).
    let eventQuestion: GoogleAppsScript.Forms.CheckboxItem | null = null;

    const cbItems = form.getItems(FormApp.ItemType.CHECKBOX);
    for (const item of cbItems) {
      try {
        if (String(item.getTitle() || '').trim() === 'Event') {
          eventQuestion = item.asCheckboxItem();
          break;
        }
      } catch {
        // ignore
      }
    }

    if (!eventQuestion) {
      // Remove legacy Event items to avoid duplicates and recreate as checkbox.
      removeItemsByTitle(form, ['Event']);
      eventQuestion = form.addCheckboxItem().setTitle('Event').setRequired(true);
    }

    // Always allow users to supply an "Other" event choice for excusals.
    try {
      // Use dynamic call because typings may not expose setOtherOption on CheckboxItem.
      const setter = (eventQuestion as any).setOtherOption || (eventQuestion as any).showOtherOption;
      if (typeof setter === 'function') setter.call(eventQuestion, true);
      else Log.warn('Unable to enable Other option on Excusals Event question: setter not available');
    } catch (err) {
      Log.warn(`Unable to enable Other option on Excusals Event question: ${err}`);
    }

    const choices: GoogleAppsScript.Forms.Choice[] = [];
    try {
      const backendId = Config.scriptProperties().getProperty(Config.PROPERTY_KEYS.BACKEND_SHEET_ID) || '';
      const eventsSheet = backendId ? SheetUtils.getSheet(backendId, 'Events Backend') : null;
      if (eventsSheet) {
        const eventsTable = SheetUtils.readTable(eventsSheet);
        const headers = (eventsTable as any).headers || [];
        if (headers.length > 0 && headers.every((h: any) => h === '' || h === null)) {
          Log.warn('Excusal events table has empty headers; check Events Backend');
        }
        eventsTable.rows.forEach((r) => {
          const name = r['display_name'] || r['attendance_column_label'] || r['event_id'];
          if (!name) return;
          choices.push(eventQuestion!.createChoice(name));
        });
      }
    } catch (err) {
      Log.warn(`Unable to populate excusals form events: ${err}`);
    }

    // Ensure at least one choice so setChoices does not fail in empty environments.
    if (choices.length === 0) {
      choices.push(eventQuestion.createChoice('No events available (add events in Events Backend)'));
    }

    eventQuestion.setChoices(choices);
    Log.info(`Excusals form: refreshed event choices count=${choices.length}`);

    enforceExcusalsItemOrder(form);
  }
}
