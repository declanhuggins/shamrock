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

  function withFormRetry<T>(
    form: GoogleAppsScript.Forms.Form,
    label: string,
    action: (f: GoogleAppsScript.Forms.Form) => T,
  ): { form: GoogleAppsScript.Forms.Form; result: T } {
    let working = form;
    let lastErr: any;
    for (let attempt = 1; attempt <= 3; attempt += 1) {
      try {
        return { form: working, result: action(working) };
      } catch (err) {
        lastErr = err;
        Log.warn(`Attendance form: ${label} failed (attempt ${attempt}): ${err}`);
        Utilities.sleep(400 * attempt);
        try {
          working = FormApp.openById(working.getId());
        } catch (reopenErr) {
          Log.warn(`Attendance form: unable to reopen form after failure: ${reopenErr}`);
        }
      }
    }
    throw lastErr;
  }

  function addPageBreakItemSafe(
    form: GoogleAppsScript.Forms.Form,
    title: string,
    goTo: GoogleAppsScript.Forms.PageNavigationType,
  ): { form: GoogleAppsScript.Forms.Form; item: GoogleAppsScript.Forms.PageBreakItem } {
    const { form: nextForm, result } = withFormRetry(form, `addPageBreakItem ${title}`, (f) => f.addPageBreakItem());
    result.setTitle(title);
    result.setGoToPage(goTo);
    return { form: nextForm, item: result };
  }

  function addTextItemSafe(
    form: GoogleAppsScript.Forms.Form,
    title: string,
    required = false,
  ): { form: GoogleAppsScript.Forms.Form; item: GoogleAppsScript.Forms.TextItem } {
    const { form: nextForm, result } = withFormRetry(form, `addTextItem ${title}`, (f) => f.addTextItem());
    result.setTitle(title);
    if (required) result.setRequired(true);
    return { form: nextForm, item: result };
  }

  function addListItemSafe(
    form: GoogleAppsScript.Forms.Form,
    title: string,
    choices?: string[],
    required = false,
  ): { form: GoogleAppsScript.Forms.Form; item: GoogleAppsScript.Forms.ListItem } {
    const { form: nextForm, result } = withFormRetry(form, `addListItem ${title}`, (f) => f.addListItem());
    result.setTitle(title);
    if (choices && choices.length) result.setChoiceValues(choices);
    if (required) result.setRequired(true);
    return { form: nextForm, item: result };
  }

  function addMultipleChoiceItemSafe(
    form: GoogleAppsScript.Forms.Form,
    title: string,
    required = false,
  ): { form: GoogleAppsScript.Forms.Form; item: GoogleAppsScript.Forms.MultipleChoiceItem } {
    const { form: nextForm, result } = withFormRetry(form, `addMultipleChoiceItem ${title}`, (f) => f.addMultipleChoiceItem());
    result.setTitle(title);
    if (required) result.setRequired(true);
    return { form: nextForm, item: result };
  }

  function addCheckboxItemSafe(
    form: GoogleAppsScript.Forms.Form,
    title: string,
    choices: string[],
  ): { form: GoogleAppsScript.Forms.Form; item: GoogleAppsScript.Forms.CheckboxItem } {
    const { form: nextForm, result } = withFormRetry(form, `addCheckboxItem ${title}`, (f) => f.addCheckboxItem());
    result.setTitle(title);
    if (choices.length > 0) {
      result.setChoiceValues(choices);
    }
    return { form: nextForm, item: result };
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
    pocByAs: Record<string, string[]>; // AS -> labels (AS300+ only)
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
    const groups: CadetGroups = { byFlight: {}, byFlightAll: {}, byCrosstown: {}, allByAs: {}, nonAbroadByAs: {}, pocByAs: {} };
    try {
      const backendId = Config.getBackendId();
      const sheet = SheetUtils.getSheet(backendId, 'Directory Backend');
      if (!sheet) return groups;
      const table = SheetUtils.readTable(sheet);
      const asYearNumber = (raw: string): number => {
        const match = String(raw || '').toUpperCase().match(/AS\s*(\d+)/);
        if (!match) return 0;
        return Number(match[1] || 0) || 0;
      };
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

        // POC Third Hour: AS300+ only, exclude Abroad cadets
        if (asYearNumber(as) >= 300 && !isAbroad) {
          groups.pocByAs[as] = groups.pocByAs[as] || [];
          groups.pocByAs[as].push(label);
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
      sortValues(groups.pocByAs);
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

    // Section 1: respondent info
    let workingForm = form;
    const nameItem = addTextItemSafe(workingForm, 'Name', true);
    workingForm = nameItem.form;

    // Section 2: Event category selection
    const eventCategoryPage = addPageBreakItemSafe(workingForm, 'Event Category', FormApp.PageNavigationType.CONTINUE);
    workingForm = eventCategoryPage.form;
    const eventCategoryItem = addMultipleChoiceItemSafe(workingForm, 'Event Type', true);
    workingForm = eventCategoryItem.form;

    // Section 3: Event list pages for each category (to be populated later)
    const mandoEventsPage = addPageBreakItemSafe(workingForm, 'Mando Events', FormApp.PageNavigationType.CONTINUE);
    workingForm = mandoEventsPage.form;
    const mandoEventList = addListItemSafe(workingForm, 'Select Event (Mando)', undefined, true);
    workingForm = mandoEventList.form;

    const llabEventsPage = addPageBreakItemSafe(workingForm, 'LLAB Events', FormApp.PageNavigationType.CONTINUE);
    workingForm = llabEventsPage.form;
    const llabEventList = addListItemSafe(workingForm, 'Select Event (LLAB)', undefined, true);
    workingForm = llabEventList.form;

    const pocEventsPage = addPageBreakItemSafe(workingForm, 'Third Hour Events', FormApp.PageNavigationType.CONTINUE);
    workingForm = pocEventsPage.form;
    const pocEventList = addListItemSafe(workingForm, 'Select Event (POC Third Hour)', undefined, true);
    workingForm = pocEventList.form;

    const secondaryEventsPage = addPageBreakItemSafe(workingForm, 'Secondary Events', FormApp.PageNavigationType.CONTINUE);
    workingForm = secondaryEventsPage.form;
    const secondaryEventList = addListItemSafe(workingForm, 'Select Event (Secondary)', undefined, true);
    workingForm = secondaryEventList.form;

    const otherEventsPage = addPageBreakItemSafe(workingForm, 'Other Events', FormApp.PageNavigationType.CONTINUE);
    workingForm = otherEventsPage.form;
    const otherEventList = addListItemSafe(workingForm, 'Select Event (Other)', undefined, true);
    workingForm = otherEventList.form;

    // Section 4: Mando branch selector
    const mandoBranch = addPageBreakItemSafe(workingForm, 'Mando Branch', FormApp.PageNavigationType.CONTINUE);
    workingForm = mandoBranch.form;
    const mandoFlight = addMultipleChoiceItemSafe(workingForm, 'Flight / Crosstown (Mando)', true);
    workingForm = mandoFlight.form;
    const mandoFlightItem = mandoFlight.item;
    const mandoFlights = [...Arrays.FLIGHTS.filter((f) => f !== 'Abroad'), 'Trine', 'Valparaiso'];
    const mandoFlightPages: Record<string, GoogleAppsScript.Forms.PageBreakItem> = {};
    mandoFlights.forEach((fName) => {
      const pageResult = addPageBreakItemSafe(workingForm, `Cadets for ${fName} (Mando)`, FormApp.PageNavigationType.SUBMIT);
      workingForm = pageResult.form;
      const page = pageResult.item;
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
          const result = addCheckboxItemSafe(workingForm, `Cadets (${fName}) AS ${as} (Mando)`, opts);
          workingForm = result.form;
        });
    });
    mandoFlightItem.setChoices(mandoFlights.map((f) => mandoFlightItem.createChoice(f, mandoFlightPages[f])));
    Log.info(`Attendance form: Mando flight pages=${mandoFlights.length}`);

    // Section 5: LLAB branch selector
    const llabBranch = addPageBreakItemSafe(workingForm, 'LLAB Branch', FormApp.PageNavigationType.CONTINUE);
    workingForm = llabBranch.form;
    const llabFlight = addMultipleChoiceItemSafe(workingForm, 'Flight (LLAB)', true);
    workingForm = llabFlight.form;
    const llabFlightItem = llabFlight.item;
    const llabFlights = Arrays.FLIGHTS.filter((f) => f !== 'Abroad');
    const llabFlightPages: Record<string, GoogleAppsScript.Forms.PageBreakItem> = {};
    llabFlights.forEach((fName) => {
      const pageResult = addPageBreakItemSafe(workingForm, `Cadets for ${fName} (LLAB)`, FormApp.PageNavigationType.SUBMIT);
      workingForm = pageResult.form;
      const page = pageResult.item;
      llabFlightPages[fName] = page;

      const groupMap = cadets.byFlightAll[fName] || cadets.byFlight[fName] || {};
      Object.keys(groupMap)
        .sort((a, b) => b.localeCompare(a, undefined, { sensitivity: 'base' }))
        .forEach((as) => {
          const opts = groupMap[as];
          if (!opts || !opts.length) return;
          const result = addCheckboxItemSafe(workingForm, `Cadets (${fName}) AS ${as} (LLAB)`, opts);
          workingForm = result.form;
        });
    });
    llabFlightItem.setChoices(llabFlights.map((f) => llabFlightItem.createChoice(f, llabFlightPages[f])));
    Log.info(`Attendance form: LLAB flight pages=${llabFlights.length}`);

    // Section 6: POC Third Hour branch (AS300+ only, excludes Abroad)
    const pocPage = addPageBreakItemSafe(workingForm, 'POC Branch', FormApp.PageNavigationType.SUBMIT);
    workingForm = pocPage.form;
    const pocStart = pocPage.item;
    Object.keys(cadets.pocByAs)
      .sort((a, b) => b.localeCompare(a, undefined, { sensitivity: 'base' }))
      .forEach((as) => {
        const opts = cadets.pocByAs[as];
        if (!opts || !opts.length) return;
        const result = addCheckboxItemSafe(workingForm, `Cadets (POC) AS ${as}`, opts);
        workingForm = result.form;
      });

    // Section 7: Secondary branch (non-abroad) -> submit
    const secondaryPage = addPageBreakItemSafe(workingForm, 'Secondary Branch', FormApp.PageNavigationType.SUBMIT);
    workingForm = secondaryPage.form;
    const secondaryStart = secondaryPage.item;
    Object.keys(cadets.nonAbroadByAs)
      .sort((a, b) => b.localeCompare(a, undefined, { sensitivity: 'base' }))
      .forEach((as) => {
        const opts = cadets.nonAbroadByAs[as];
        if (!opts || !opts.length) return;
        const result = addCheckboxItemSafe(workingForm, `Cadets (Secondary) AS ${as}`, opts);
        workingForm = result.form;
      });

    // Section 8: Other branch (all cadets) -> submit
    const fallbackPage = addPageBreakItemSafe(workingForm, 'Attendance Branch', FormApp.PageNavigationType.SUBMIT);
    workingForm = fallbackPage.form;
    const fallbackStart = fallbackPage.item;
    Object.keys(cadets.allByAs)
      .sort((a, b) => b.localeCompare(a, undefined, { sensitivity: 'base' }))
      .forEach((as) => {
        const opts = cadets.allByAs[as];
        if (!opts || !opts.length) return;
        const result = addCheckboxItemSafe(workingForm, `Cadets (All) AS ${as}`, opts);
        workingForm = result.form;
      });

    // Populate event lists per category and wire navigation
    const backendId = Config.getBackendId();
    const eventsSheet = backendId ? SheetUtils.getSheet(backendId, 'Events Backend') : null;

    const mandoEventChoices: GoogleAppsScript.Forms.Choice[] = [];
    const llabEventChoices: GoogleAppsScript.Forms.Choice[] = [];
    const pocEventChoices: GoogleAppsScript.Forms.Choice[] = [];
    const secondaryEventChoices: GoogleAppsScript.Forms.Choice[] = [];
    const otherEventChoices: GoogleAppsScript.Forms.Choice[] = [];

    if (eventsSheet) {
      const eventsTable = SheetUtils.readTable(eventsSheet);
      eventsTable.rows.forEach((r) => {
        const name = r['display_name'] || r['attendance_column_label'] || r['event_id'];
        if (!name) return;
        const type = String(r['event_type'] || '').toLowerCase();
        const expectedGroup = String(r['expected_group'] || '').toLowerCase();
        const nameLc = String(name || '').toLowerCase();

        // Categorize event and determine target branch
        if (type.includes('llab')) {
          llabEventChoices.push(llabEventList.item.createChoice(name, llabBranch.item));
        } else if (type.includes('mando')) {
          mandoEventChoices.push(mandoEventList.item.createChoice(name, mandoBranch.item));
        } else if (type.includes('secondary')) {
          secondaryEventChoices.push(secondaryEventList.item.createChoice(name, secondaryPage.item));
        } else if (expectedGroup.includes('poc') || type.includes('third hour') || nameLc.includes('poc third hour')) {
          pocEventChoices.push(pocEventList.item.createChoice(name, pocPage.item));
        } else {
          otherEventChoices.push(otherEventList.item.createChoice(name, fallbackPage.item));
        }
      });
    }

    // Set choices for each event list (with fallback if empty)
    if (mandoEventChoices.length) mandoEventList.item.setChoices(mandoEventChoices);
    else mandoEventList.item.setChoices([mandoEventList.item.createChoice('(no events)', mandoBranch.item)]);

    if (llabEventChoices.length) llabEventList.item.setChoices(llabEventChoices);
    else llabEventList.item.setChoices([llabEventList.item.createChoice('(no events)', llabBranch.item)]);

    if (pocEventChoices.length) pocEventList.item.setChoices(pocEventChoices);
    else pocEventList.item.setChoices([pocEventList.item.createChoice('(no events)', pocPage.item)]);

    if (secondaryEventChoices.length) secondaryEventList.item.setChoices(secondaryEventChoices);
    else secondaryEventList.item.setChoices([secondaryEventList.item.createChoice('(no events)', secondaryPage.item)]);

    if (otherEventChoices.length) otherEventList.item.setChoices(otherEventChoices);
    else otherEventList.item.setChoices([otherEventList.item.createChoice('(no events)', fallbackPage.item)]);

    // Wire event category chooser to event list pages
    const categoryChoices: GoogleAppsScript.Forms.Choice[] = [];
    if (mandoEventChoices.length) categoryChoices.push(eventCategoryItem.item.createChoice('Mando PT', mandoEventsPage.item));
    if (llabEventChoices.length) categoryChoices.push(eventCategoryItem.item.createChoice('LLAB', llabEventsPage.item));
    if (pocEventChoices.length) categoryChoices.push(eventCategoryItem.item.createChoice('POC Third Hour', pocEventsPage.item));
    if (secondaryEventChoices.length) categoryChoices.push(eventCategoryItem.item.createChoice('Secondary', secondaryEventsPage.item));
    if (otherEventChoices.length) categoryChoices.push(eventCategoryItem.item.createChoice('Other', otherEventsPage.item));

    if (categoryChoices.length) {
      eventCategoryItem.item.setChoices(categoryChoices);
      Log.info(
        `Attendance form: event categories wired mando=${mandoEventChoices.length} llab=${llabEventChoices.length} poc=${pocEventChoices.length} secondary=${secondaryEventChoices.length} other=${otherEventChoices.length}`
      );
    } else {
      eventCategoryItem.item.setChoices([
        eventCategoryItem.item.createChoice('(set up events first)', fallbackPage.item),
      ]);
      Log.warn('Attendance form: no events found to populate categories');
    }

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
    const pocStart = findPage('POC Branch');
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
          const expectedGroup = String(r['expected_group'] || '').toLowerCase();
          const nameLc = String(name || '').toLowerCase();
          let target: GoogleAppsScript.Forms.PageBreakItem | GoogleAppsScript.Forms.PageNavigationType | null = fallbackStart;
          if (type.includes('llab')) target = llabStart;
          else if (type.includes('mando')) target = mandoStart;
          else if (type.includes('secondary')) target = secondaryStart;
          else if (expectedGroup.includes('poc') || type.includes('third hour') || nameLc.includes('poc third hour')) target = pocStart;

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
    // Rebuild form with structured flow: Name → Event Type → Events → Attendance Type → Reason
    rebuildExcusalsForm(form);
  }

  export function rebuildExcusalsForm(form: GoogleAppsScript.Forms.Form) {
    Log.info('Excusals form: start rebuild');
    clearItems(form);
    Log.info('Excusals form: cleared existing items');

    // Section 1: Respondent info
    let workingForm = form;
    const lastNameItem = addTextItemSafe(workingForm, 'Last Name', true);
    workingForm = lastNameItem.form;
    const firstNameItem = addTextItemSafe(workingForm, 'First Name', true);
    workingForm = firstNameItem.form;

    // Section 2: Event category selection (with loop-back navigation)
    const eventCategoryPage = addPageBreakItemSafe(workingForm, 'Event Category', FormApp.PageNavigationType.CONTINUE);
    workingForm = eventCategoryPage.form;
    const eventCategoryItem = addMultipleChoiceItemSafe(workingForm, 'Select Event Type (or Done to continue)', true);
    workingForm = eventCategoryItem.form;

    // Section 3: Event selection pages per category (checkboxes to allow multiple)
    // Each page continues back to Event Category
    const mandoEventsPage = addPageBreakItemSafe(workingForm, 'Mando Events', FormApp.PageNavigationType.CONTINUE);
    workingForm = mandoEventsPage.form;
    const mandoEventList = addCheckboxItemSafe(workingForm, 'Select Event(s) (Mando)', []);
    workingForm = mandoEventList.form;

    const llabEventsPage = addPageBreakItemSafe(workingForm, 'LLAB Events', FormApp.PageNavigationType.CONTINUE);
    workingForm = llabEventsPage.form;
    const llabEventList = addCheckboxItemSafe(workingForm, 'Select Event(s) (LLAB)', []);
    workingForm = llabEventList.form;

    const pocEventsPage = addPageBreakItemSafe(workingForm, 'Third Hour Events', FormApp.PageNavigationType.CONTINUE);
    workingForm = pocEventsPage.form;
    const pocEventList = addCheckboxItemSafe(workingForm, 'Select Event(s) (POC Third Hour)', []);
    workingForm = pocEventList.form;

    const secondaryEventsPage = addPageBreakItemSafe(workingForm, 'Secondary Events', FormApp.PageNavigationType.CONTINUE);
    workingForm = secondaryEventsPage.form;
    const secondaryEventList = addCheckboxItemSafe(workingForm, 'Select Event(s) (Secondary)', []);
    workingForm = secondaryEventList.form;

    const otherEventsPage = addPageBreakItemSafe(workingForm, 'Other Events', FormApp.PageNavigationType.CONTINUE);
    workingForm = otherEventsPage.form;
    const otherEventList = addCheckboxItemSafe(workingForm, 'Select Event(s) (Other)', []);
    workingForm = otherEventList.form;

    // Section 4: Attendance type and reason (shared final section)
    const detailsPage = addPageBreakItemSafe(workingForm, 'Excusal Details', FormApp.PageNavigationType.SUBMIT);
    workingForm = detailsPage.form;
    const attendanceTypeItem = addListItemSafe(workingForm, 'Requested Attendance Type', ['E', 'ES', 'MU', 'MRS'], true);
    workingForm = attendanceTypeItem.form;
    const reasonItem = addTextItemSafe(workingForm, 'Reason', true);
    workingForm = reasonItem.form;

    // Set each event page to navigate back to Event Category after event selection
    mandoEventsPage.item.setGoToPage(eventCategoryPage.item);
    llabEventsPage.item.setGoToPage(eventCategoryPage.item);
    pocEventsPage.item.setGoToPage(eventCategoryPage.item);
    secondaryEventsPage.item.setGoToPage(eventCategoryPage.item);
    otherEventsPage.item.setGoToPage(eventCategoryPage.item);

    // Populate event lists per category
    const backendId = Config.getBackendId();
    const eventsSheet = backendId ? SheetUtils.getSheet(backendId, 'Events Backend') : null;

    const mandoEventChoices: string[] = [];
    const llabEventChoices: string[] = [];
    const pocEventChoices: string[] = [];
    const secondaryEventChoices: string[] = [];
    const otherEventChoices: string[] = [];

    if (eventsSheet) {
      const eventsTable = SheetUtils.readTable(eventsSheet);
      eventsTable.rows.forEach((r) => {
        const name = r['display_name'] || r['attendance_column_label'] || r['event_id'];
        if (!name) return;
        const type = String(r['event_type'] || '').toLowerCase();
        const expectedGroup = String(r['expected_group'] || '').toLowerCase();
        const nameLc = String(name || '').toLowerCase();

        if (type.includes('llab')) {
          llabEventChoices.push(name);
        } else if (type.includes('mando')) {
          mandoEventChoices.push(name);
        } else if (type.includes('secondary')) {
          secondaryEventChoices.push(name);
        } else if (expectedGroup.includes('poc') || type.includes('third hour') || nameLc.includes('poc third hour')) {
          pocEventChoices.push(name);
        } else {
          otherEventChoices.push(name);
        }
      });
    }

    // Set choices for each event list
    if (mandoEventChoices.length) {
      const choices = mandoEventChoices.map(c => mandoEventList.item.createChoice(c));
      mandoEventList.item.setChoices(choices);
    } else {
      mandoEventList.item.setChoices([mandoEventList.item.createChoice('(no events)')]);
    }

    if (llabEventChoices.length) {
      const choices = llabEventChoices.map(c => llabEventList.item.createChoice(c));
      llabEventList.item.setChoices(choices);
    } else {
      llabEventList.item.setChoices([llabEventList.item.createChoice('(no events)')]);
    }

    if (pocEventChoices.length) {
      const choices = pocEventChoices.map(c => pocEventList.item.createChoice(c));
      pocEventList.item.setChoices(choices);
    } else {
      pocEventList.item.setChoices([pocEventList.item.createChoice('(no events)')]);
    }

    if (secondaryEventChoices.length) {
      const choices = secondaryEventChoices.map(c => secondaryEventList.item.createChoice(c));
      secondaryEventList.item.setChoices(choices);
    } else {
      secondaryEventList.item.setChoices([secondaryEventList.item.createChoice('(no events)')]);
    }

    if (otherEventChoices.length) {
      const choices = otherEventChoices.map(c => otherEventList.item.createChoice(c));
      otherEventList.item.setChoices(choices);
    } else {
      otherEventList.item.setChoices([otherEventList.item.createChoice('(no events)')]);
    }

    // Wire event category chooser to event list pages
    const categoryChoices: GoogleAppsScript.Forms.Choice[] = [];
    if (mandoEventChoices.length) categoryChoices.push(eventCategoryItem.item.createChoice('Mando PT', mandoEventsPage.item));
    if (llabEventChoices.length) categoryChoices.push(eventCategoryItem.item.createChoice('LLAB', llabEventsPage.item));
    if (pocEventChoices.length) categoryChoices.push(eventCategoryItem.item.createChoice('POC Third Hour', pocEventsPage.item));
    if (secondaryEventChoices.length) categoryChoices.push(eventCategoryItem.item.createChoice('Secondary', secondaryEventsPage.item));
    if (otherEventChoices.length) categoryChoices.push(eventCategoryItem.item.createChoice('Other', otherEventsPage.item));
    
    // Add "Done" option to proceed to details page
    categoryChoices.push(eventCategoryItem.item.createChoice('Done selecting events', detailsPage.item));

    if (categoryChoices.length > 1) {
      eventCategoryItem.item.setChoices(categoryChoices);
      Log.info(
        `Excusals form: event categories wired mando=${mandoEventChoices.length} llab=${llabEventChoices.length} poc=${pocEventChoices.length} secondary=${secondaryEventChoices.length} other=${otherEventChoices.length}`
      );
    } else {
      eventCategoryItem.item.setChoices([
        eventCategoryItem.item.createChoice('Done selecting events', detailsPage.item),
      ]);
      Log.warn('Excusals form: no events found to populate categories');
    }

    Log.info('Excusals form: completed rebuild');
  }

  export function refreshExcusalsFormEventChoices(form: GoogleAppsScript.Forms.Form) {
    // Legacy function - now handled by rebuildExcusalsForm
    Log.info('refreshExcusalsFormEventChoices: use rebuildExcusalsForm instead');
  }
}
