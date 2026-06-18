export function attachDirtyTracker(instance, { onDirty, onSignal } = {}) {
    let dirty = false;
    let armed = false;
    const unsubs = [];

    function listen(target, event, handler) {
        if (!target?.addEventListener) return;
        target.addEventListener(event, handler);
        unsubs.push(() => target.removeEventListener?.(event, handler));
    }

    function markDirty(source) {
        if (dirty) return;
        if (!armed) {
            onSignal?.(`ignored (pre-load): ${source}`);
            return;
        }
        dirty = true;
        onDirty?.(source);
    }

    function handlePagesUpdated(payload) {
        const changes = payload && typeof payload === 'object' ? payload : {};
        const added = changes.added?.length ?? 0;
        const removed = changes.removed?.length ?? 0;
        const rotated = changes.rotationChanged?.length ?? 0;
        const moved = changes.moved && typeof changes.moved === 'object'
            ? Object.keys(changes.moved).length
            : 0;
        const detail = {
            added, removed, rotated, moved,
            contentChanged: changes.contentChanged?.length ?? 0,
            linearizedUpdate: changes.linearizedUpdate ?? false,
        };
        onSignal?.('pagesUpdated', detail);
        if (added > 0 || removed > 0 || rotated > 0 || moved > 0) {
            markDirty('pagesUpdated');
        }
    }

    const { Core, UI } = instance;
    const documentViewer = Core?.documentViewer;
    const annotationManager = Core?.annotationManager;

    function arm() {
        if (armed) return;
        try {
            documentViewer?.getAnnotationHistoryManager?.()?.clear?.();
        } catch (e) { console.log(e); }
        armed = true;
        onSignal?.('armed (document fully loaded)');
    }

    listen(documentViewer, 'annotationsLoaded', arm);
    const loadedPromise = documentViewer?.getAnnotationsLoadedPromise?.();
    if (loadedPromise) {
        onSignal?.('awaiting getAnnotationsLoadedPromise() before wiring');
        Promise.resolve(loadedPromise).then(arm).catch(arm);
    }

    function listenHistory(manager, event, source) {
        if (!manager) return;
        listen(manager, event, () => {
            const canUndo = manager.canUndo?.() ?? false;
            onSignal?.(source, { event, canUndo });
            if (canUndo) markDirty(source);
        });
    }

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

    listen(documentViewer, 'pagesUpdated', (...args) => handlePagesUpdated(args[0]));

    listen(annotationManager, 'fieldChanged', () => {
        onSignal?.('fieldChanged');
        markDirty('fieldChanged');
    });

    listen(annotationManager, 'annotationChanged', (...args) => {
        const annotations = args[0];
        const info = args[2];
        const imported = info?.imported ?? false;
        const isForm = Array.isArray(annotations) && annotations.some(a =>
            a?.elementName === 'widget' ||
            a?.Subject === 'Widget' ||
            typeof a?.getField === 'function' ||
            a?.fieldName != null
        );
        onSignal?.('annotationChanged', { imported, isForm });
        if (!imported && isForm) markDirty('annotationChanged');
    });

    listen(UI, 'outlineBookmarksChanged', () => {
        onSignal?.('outlineBookmarksChanged');
        markDirty('outlineBookmarksChanged');
    });

    listen(documentViewer, 'documentLoaded', () => {
        try {
            const doc = documentViewer?.getDocument?.();
            const officeEditor = doc?.getOfficeEditor?.();

            if (doc?.addEventListener) {
                listen(doc, 'officeDocumentEdited', () => {
                    onSignal?.('officeDocumentEdited (doc)');
                    markDirty('officeDocumentEdited');
                });
            }

            if (officeEditor?.addEventListener) {
                listen(officeEditor, 'officeDocumentEdited', () => {
                    onSignal?.('officeDocumentEdited (editor)');
                    markDirty('officeDocumentEdited');
                });
            }
        } catch (e) {
            onSignal?.('officeEditor not available (expected for spreadsheets)');
        }

        const SpreadsheetEditor = Core?.SpreadsheetEditor;
        const SEEvents = SpreadsheetEditor?.SpreadsheetEditorManager?.Events;
        const spreadsheetManager = documentViewer?.getSpreadsheetEditorManager?.();

        onSignal?.('spreadsheet check', {
            hasManager: !!spreadsheetManager,
            hasEvents: !!SEEvents,
            eventKeys: SEEvents ? Object.keys(SEEvents) : [],
        });

        if (spreadsheetManager && SEEvents) {
            const readyEvent = SEEvents.SPREADSHEET_EDITOR_READY;
            if (readyEvent) {
                spreadsheetManager.addEventListener(readyEvent, () => {
                    onSignal?.('SPREADSHEET_EDITOR_READY fired');

                    const skipEvents = new Set([
                        'SPREADSHEET_EDITOR_READY',
                        'SPREADSHEET_EDITOR_LOADED',
                        'SELECTION_CHANGED',
                        'ACTIVE_SHEET_CHANGED',
                        'FORMULA_BAR_TEXT_CHANGED',
                    ]);

                    for (const [name, value] of Object.entries(SEEvents)) {
                        if (skipEvents.has(name)) continue;
                        spreadsheetManager.addEventListener(value, () => {
                            onSignal?.(`spreadsheet: ${name}`);
                            markDirty(`spreadsheet:${name}`);
                        });
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

    return {
        isDirty: () => dirty,
        reset: () => { dirty = false; },
        dispose: () => {
            for (const unsub of unsubs.splice(0)) {
                try { unsub(); } catch (e) { console.log(e) }
            }
        },
    };
}