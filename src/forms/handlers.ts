// Form submission handlers: send receipts and (future) processing hooks.

namespace FormHandlers {
	function getNamedValues(e: GoogleAppsScript.Events.FormsOnFormSubmit): Record<string, string[]> {
		return ((e as any).namedValues as Record<string, string[]>) || {};
	}

	function getFirstNamedValue(namedValues: Record<string, string[]>, key: string): string {
		const raw = namedValues[key];
		if (!raw) return '';
		const arr = Array.isArray(raw) ? raw : [raw];
		return String(arr[0] || '').trim();
	}

	function findFirstValue(namedValues: Record<string, string[]>, matcher: (key: string) => boolean): string {
		for (const [k, vals] of Object.entries(namedValues)) {
			if (!matcher(k)) continue;
			const arr = Array.isArray(vals) ? vals : [vals];
			const first = String(arr[0] || '').trim();
			if (first) return first;
		}
		return '';
	}

	function formatTimestamp(dateLike: any): string {
		try {
			const d = dateLike instanceof Date ? dateLike : new Date(dateLike);
			return d.toISOString();
		} catch (err) {
			return new Date().toISOString();
		}
	}

	function appendToBackend(sheetName: string, rows: Record<string, any>[]) {
		SheetUtils.appendRows(Config.getBackendSheet(sheetName), rows);
	}

	function lookupCadetByEmail(email: string) {
		const backendId = Config.getBackendId();
		if (!backendId || !email) return null;
		const directorySheet = SheetUtils.getSheet(backendId, 'Directory Backend');
		if (!directorySheet) return null;
		const data = SheetUtils.readTable(directorySheet);
		const lower = email.toLowerCase();
		return data.rows.find((r) => String(r['email'] || '').toLowerCase() === lower) || null;
	}

	function sendReceipt(opts: { to: string; subject: string; body: string; replyTo?: string }) {
		if (!opts.to) return;
		GmailApp.sendEmail(opts.to, opts.subject, opts.body, {
			replyTo: opts.replyTo,
			name: 'SHAMROCK Automations',
		});
	}

	function summarizeResponses(e: GoogleAppsScript.Events.FormsOnFormSubmit): string {
		try {
			const resp = e.response;
			const items = resp.getItemResponses();
			return items
				.map((ir) => `${ir.getItem().getTitle()}: ${ir.getResponse()}`)
				.join('\n');
		} catch (err) {
			return 'Summary unavailable.';
		}
	}

	function getEmail(e: GoogleAppsScript.Events.FormsOnFormSubmit): string {
		const fromResponse = e.response.getRespondentEmail();
		if (fromResponse) return String(fromResponse).trim();
		const nv = (e as any).namedValues || {};
		const keys = ['Email'];
		for (const k of keys) {
			const v = nv[k];
			if (v) {
				const s = Array.isArray(v) ? v[0] : v;
				if (s) return String(s).trim();
			}
		}
		return '';
	}

	function getEditUrl(e: GoogleAppsScript.Events.FormsOnFormSubmit): string {
		try {
			return e.response.getEditResponseUrl();
		} catch (err) {
			return '';
		}
	}

	export function onDirectoryFormSubmit(e: GoogleAppsScript.Events.FormsOnFormSubmit) {
		DirectoryService.handleDirectoryFormSubmission(e);
	}

		export function onAttendanceFormSubmit(e: GoogleAppsScript.Events.FormsOnFormSubmit) {
			const namedValues = getNamedValues(e);
			const submittedByEmail = getEmail(e);
			const submittedAt = formatTimestamp(e.response?.getTimestamp?.() || new Date());
			const submissionId = `att-${Date.now()}`;

			// Use item responses for robustness (titles can vary), with namedValues as fallback.
			let eventName = '';
			let submittedByName = '';
			let flight = '';
			const cadetSelections: string[] = [];

			const normalizeToList = (val: any): string[] => {
				if (Array.isArray(val)) {
					return val
						.map((v) => String(v || ''))
						.flatMap((v) => v.split(',').map((s) => s.trim()))
						.filter(Boolean);
				}
				const s = String(val || '').trim();
				if (!s) return [];
				return s
					.split(',')
					.map((x) => x.trim())
					.filter(Boolean);
			};

			try {
				e.response.getItemResponses().forEach((ir) => {
					const title = (ir.getItem().getTitle() || '').trim();
					const lower = title.toLowerCase();
					const resp = ir.getResponse();
					const list = normalizeToList(resp);

					if (!eventName && lower === 'event') {
						eventName = list[0] || String(resp || '').trim();
					}

					if (!submittedByName && lower.includes('name') && !lower.includes('cadets')) {
						submittedByName = list[0] || String(resp || '').trim();
					}

					if (!flight && lower.includes('flight')) {
						flight = list[0] || String(resp || '').trim();
					}

					if (lower.includes('cadets')) {
						list.forEach((n) => cadetSelections.push(n));
					}
				});
			} catch (err) {
				Log.warn(`attendance parse (items) failed: ${err}`);
			}

			// Fallbacks to namedValues if still empty.
			if (!eventName) {
				eventName =
					getFirstNamedValue(namedValues, 'Event') ||
					findFirstValue(namedValues, (k) => k.toLowerCase().trim() === 'event');
			}
			if (!submittedByName) {
				submittedByName =
					getFirstNamedValue(namedValues, 'Name') ||
					getFirstNamedValue(namedValues, 'Submitted By Name') ||
					findFirstValue(namedValues, (k) => k.toLowerCase().includes('name'));
			}
			if (!flight) {
				flight =
					getFirstNamedValue(namedValues, 'Flight / Crosstown (LLAB)') ||
					getFirstNamedValue(namedValues, 'Flight (Mando PT)') ||
					getFirstNamedValue(namedValues, 'Flight / Crosstown') ||
					getFirstNamedValue(namedValues, 'Flight (LLAB)') ||
					findFirstValue(namedValues, (k) => k.toLowerCase().includes('flight'));
			}

			if (!cadetSelections.length) {
				Object.entries(namedValues).forEach(([key, vals]) => {
					const lowerKey = key.toLowerCase();
					if (!lowerKey.includes('cadets')) return;
					const list = normalizeToList(vals);
					list.forEach((n) => cadetSelections.push(n));
				});
			}

			// If responses are alternating Last/First, re-pair into "Last, First" before dedupe.
			const normalizeCadetNames = (names: string[]): string[] => {
				if (!names.length) return names;
				const hasComma = names.some((n) => n.includes(','));
				if (hasComma) return names;
				if (names.length < 2) return names;
				const paired: string[] = [];
				for (let i = 0; i < names.length; i += 2) {
					const last = names[i];
					const first = names[i + 1];
					if (first !== undefined && first !== '') {
						paired.push(`${last}, ${first}`.trim());
					} else {
						paired.push(last);
					}
				}
				return paired;
			};

			const normalizedCadets = normalizeCadetNames(cadetSelections);

			// Deduplicate while preserving first-seen order.
			const seen = new Set<string>();
			const cadetField = normalizedCadets
				.filter((n) => {
					const key = n.toLowerCase();
					if (seen.has(key)) return false;
					seen.add(key);
					return true;
				})
				.join('; ');

			const flightFromDirectory = submittedByEmail ? lookupCadetByEmail(submittedByEmail)?.flight || '' : '';

			appendToBackend('Attendance Backend', [
				{
					submission_id: submissionId,
					submitted_at: submittedAt,
					event: eventName,
					attendance_type: 'P',
					email: submittedByEmail,
					name: submittedByName,
					flight: flight || flightFromDirectory,
					cadets: cadetField,
				},
			]);

			// Incrementally apply the new log entry to the matrices (no full rebuild).
			AttendanceService.applyAttendanceLogEntry({
				event: eventName,
				attendance_type: 'P',
				cadets: cadetField,
			});

			SetupService.applyAttendanceBackendFormattingPublic();

			// Deliberately omit email receipt for attendance submissions per policy.
		}

	export function onExcusalsFormSubmit(e: GoogleAppsScript.Events.FormsOnFormSubmit) {
		const namedValues = getNamedValues(e);
		const email = getEmail(e);
		const submittedAt = formatTimestamp(e.response?.getTimestamp?.() || new Date());

		// Look up cadet info from Directory Backend
		const cadet = lookupCadetByEmail(email);

		// Parse form responses using item responses for robustness
		let events: string[] = [];
		let lastName = '';
		let firstName = '';
		let reason = '';
		const itemResponses = e.response?.getItemResponses() || [];
		
		for (const itemResponse of itemResponses) {
			const title = itemResponse.getItem().getTitle();
			const response = itemResponse.getResponse();
			
			if (title === 'Event') {
				if (Array.isArray(response)) {
					events = response.map((e) => String(e || '').trim()).filter(Boolean);
				} else {
					const eventRaw = String(response || '').trim();
					events = eventRaw
						.split(',')
						.map((ev) => ev.trim())
						.filter(Boolean);
				}
			} else if (title === 'Last Name') {
				lastName = String(response || '').trim();
			} else if (title === 'First Name') {
				firstName = String(response || '').trim();
			} else if (title === 'Reason') {
				reason = String(response || '').trim();
			}
		}

		if (events.length === 0) {
			Log.warn(`Excusal submission from ${email} has no events; skipping backend append.`);
			return;
		}

		// Create one row per event in Excusals Backend
		const rows = events.map((eventName) => {
			const requestId = `exc-${Date.now()}-${Math.random().toString(36).substring(2, 9)}`;
			return {
				request_id: requestId,
				event: eventName,
				email,
				last_name: cadet?.last_name || lastName,
				first_name: cadet?.first_name || firstName,
				flight: cadet?.flight || '',
				squadron: cadet?.squadron || '',
				status: 'Pending',
				decision: '',
				decided_by: '',
				decided_at: '',
				attendance_effect: '',
				submitted_at: submittedAt,
				last_updated_at: submittedAt,
				notes: reason,
			};
		});

		appendToBackend('Excusals Backend', rows);

		// Send notifications and sync to management panel for each row
		rows.forEach((row) => {
			ExcusalsService.notifySquadronCommanderOfNewExcusal(row);
			ExcusalsService.syncExcusalToManagementPanel(row);
			// Update attendance matrix: empty -> ER, unexcused -> UR
			ExcusalsService.updateAttendanceOnExcusalSubmission(row);
		});

		Log.info(`Excusal submission processed: ${email} requesting excusal for ${events.length} event(s)`);

		// Deliberately omit email receipt for excusal submissions per policy.
	}
}