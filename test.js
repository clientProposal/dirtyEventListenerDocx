
// 1. State and Utilities
// Tracks dirty flag, armed state, and page settle status. Provides helper
// functions for subscribing to events and marking the document as dirty.
export function attachDirtyTracker(instance, { onDirty, onSignal } = {}) {
    let dirty = false;
    let armed = false;
    let pagesSettled = false;
    let settleTimer = null;
    const SETTLE_MS = 2000;
    const unsubs = [];

    function listen(target, event, handler) {
        if (!target?.addEventListener) return;
        target.addEventListener(event, handler);
        unsubs.push(() => target.removeEventListener?.(event, handler));
    }

    function markDirty(source) {
        if (dirty) return;
        if (!armed) return;
        dirty = true;
        onDirty?.(source);
    }

    // 2. Office Annotation Filter
    // Identifies annotations created internally by the Office Editor for
    // tracked changes and comments, so they can be excluded from dirty checks.
    function isOfficeEditorAnnotation(annotation) {
        if (!annotation) return false;
        try {
            const customData = annotation.getCustomData?.('OFFICE_EDITOR_TRACKED_CHANGE_KEY')
                || annotation.getCustomData?.('OFFICE_EDITOR_COMMENT_KEY');
            if (customData) return true;
        } catch (e) { console.log(e); }

        const subject = annotation.Subject || '';
        if (subject.includes('TrackedChange') || subject.includes('Comment')) return true;

        return false;
    }

    // 3. Page Update Handler
    // Ignores pagesUpdated events during initial document rendering by waiting
    // for a settle period. Only flags dirty after all pages have loaded.
    function handlePagesUpdated(payload) {
        const changes = payload && typeof payload === 'object' ? payload : {};
        const added = changes.added?.length ?? 0;
        const removed = changes.removed?.length ?? 0;
        const rotated = changes.rotationChanged?.length ?? 0;
        const moved = changes.moved && typeof changes.moved === 'object'
            ? Object.keys(changes.moved).length
            : 0;

        if (!pagesSettled) {
            if (settleTimer) clearTimeout(settleTimer);
            settleTimer = setTimeout(() => { pagesSettled = true; }, SETTLE_MS);
            return;
        }

        if (added > 0 || removed > 0 || rotated > 0 || moved > 0) {
            markDirty('pagesUpdated');
        }
    }

    const { Core, UI } = instance;
    const documentViewer = Core?.documentViewer;
    const annotationManager = Core?.annotationManager;

    // 4. Arming Mechanism
    // Clears the annotation history and enables dirty tracking. Nothing is
    // flagged as dirty until arm() has been called after initial load settles.
    function arm() {
        if (armed) return;
        try {
            documentViewer?.getAnnotationHistoryManager?.()?.clear?.();
        } catch (e) { console.log(e); }
        armed = true;
    }

    // 5. History Manager Listeners
    // Watches annotation and content edit history managers. Only marks dirty
    // when there is an undoable action, filtering out initial load noise.
    function listenHistory(manager, event, source) {
        if (!manager) return;
        listen(manager, event, () => {
            if (manager.canUndo?.()) markDirty(source);
        });
    }

    // 6. Core Event Listeners
    // Subscribes to pagesUpdated, fieldChanged, annotationChanged, and
    // outlineBookmarksChanged. Filters out imported and Office Editor annotations.
    listen(documentViewer, 'pagesUpdated', (...args) => handlePagesUpdated(args[0]));

    listen(annotationManager, 'fieldChanged', () => markDirty('fieldChanged'));

    listen(annotationManager, 'annotationChanged', (...args) => {
        const annotations = args[0];
        const info = args[2];
        const imported = info?.imported ?? false;

        const userAnnotations = Array.isArray(annotations)
            ? annotations.filter(a => !isOfficeEditorAnnotation(a))
            : annotations;

        if (!userAnnotations || userAnnotations.length === 0) return;

        const isForm = Array.isArray(userAnnotations) && userAnnotations.some(a =>
            a?.elementName === 'widget' ||
            a?.Subject === 'Widget' ||
            typeof a?.getField === 'function' ||
            a?.fieldName != null
        );

        if (!imported && isForm) markDirty('annotationChanged');
    });

    listen(UI, 'outlineBookmarksChanged', () => markDirty('outlineBookmarksChanged'));

    // 7. Document Loaded Handler
    // Resets state on each new document. Uses a settle timer to defer arming
    // until all initial page rendering and annotation syncing has completed.
    listen(documentViewer, 'documentLoaded', () => {
        pagesSettled = false;
        armed = false;
        if (settleTimer) clearTimeout(settleTimer);

        const initialPageCount = documentViewer?.getPageCount?.() ?? 0;

        const armAfterSettle = () => {
            pagesSettled = true;
            try {
                documentViewer?.getAnnotationHistoryManager?.()?.clear?.();
            } catch (e) { console.log(e); }
            arm();
        };

        if (initialPageCount > 0) {
            settleTimer = setTimeout(armAfterSettle, SETTLE_MS);
        }

        // 8. History Wiring Post-Load
        // Attaches history listeners after documentLoaded to get fresh manager
        // references. Covers both annotation history and content edit history.
        listenHistory(
            documentViewer?.getAnnotationHistoryManager?.(),
            'historyChanged',
            'annotationHistory'
        );

        listenHistory(
            documentViewer?.getContentEditHistoryManager?.(),
            'undoRedoStatusChanged',
            'contentEdit'
        );

        // 9. Office Editor Listener
        // Listens for officeDocumentEdited on both the document and the Office
        // Editor instance. Wrapped in try-catch as it throws for spreadsheets.
        try {
            const doc = documentViewer?.getDocument?.();
            const officeEditor = doc?.getOfficeEditor?.();

            if (doc?.addEventListener) {
                listen(doc, 'officeDocumentEdited', () => markDirty('officeDocumentEdited'));
            }
            if (officeEditor?.addEventListener) {
                listen(officeEditor, 'officeDocumentEdited', () => markDirty('officeDocumentEdited'));
            }
        } catch (e) { console.log(e); }

        // 10. Spreadsheet Editor Listener
        // Waits for SPREADSHEET_EDITOR_READY then wires all edit-related events
        // from the Events enum, skipping non-edit events that fire during load.
        const SpreadsheetEditor = Core?.SpreadsheetEditor;
        const SEEvents = SpreadsheetEditor?.SpreadsheetEditorManager?.Events;
        const spreadsheetManager = documentViewer?.getSpreadsheetEditorManager?.();

        if (spreadsheetManager && SEEvents) {
            const readyEvent = SEEvents.SPREADSHEET_EDITOR_READY;
            if (readyEvent) {
                spreadsheetManager.addEventListener(readyEvent, () => {
                    const skipEvents = new Set([
                        'SPREADSHEET_EDITOR_READY',
                        'SPREADSHEET_EDITOR_LOADED',
                        'SELECTION_CHANGED',
                        'ACTIVE_SHEET_CHANGED',
                        'FORMULA_BAR_TEXT_CHANGED',
                    ]);

                    for (const [name, value] of Object.entries(SEEvents)) {
                        if (skipEvents.has(name)) continue;
                        spreadsheetManager.addEventListener(value, () => markDirty(`spreadsheet:${name}`));
                        unsubs.push(() => {
                            try { spreadsheetManager.removeEventListener?.(value); } catch (e) { console.log(e); }
                        });
                    }
                });
                unsubs.push(() => {
                    try { spreadsheetManager.removeEventListener?.(readyEvent); } catch (e) { console.log(e); }
                });
            }
        }
    });

    // 11. Deferred Page Settle
    // Second pagesUpdated listener for Office Editor documents where page count
    // is zero at documentLoaded. Arms the tracker once page events stop arriving.
    listen(documentViewer, 'pagesUpdated', (...args) => {
        if (!pagesSettled && !armed) {
            if (settleTimer) clearTimeout(settleTimer);
            settleTimer = setTimeout(() => {
                pagesSettled = true;
                arm();
            }, SETTLE_MS);
        }
    });

    // 12. Public API
    // Returns isDirty to check state, reset to clear after saving, and dispose
    // to remove all event listeners and clear timers on unmount.
    return {
        isDirty: () => dirty,
        reset: () => { dirty = false; },
        dispose: () => {
            if (settleTimer) clearTimeout(settleTimer);
            for (const unsub of unsubs.splice(0)) {
                try { unsub(); } catch (e) { console.log(e); }
            }
        },
    };
}
