import { useRef, useEffect, useState } from 'react'
import WebViewer from '@pdftron/webviewer';
import { attachDirtyTracker } from './dirtyEventListener';

function App() {
  const viewer = useRef(null);

  const inst = useRef(null);

  const hasBeenInitialized = useRef(false);
  const [dirty, setDirty] = useState(false);

  const { VITE_PDFTRONKEY: licenseKey } = import.meta.env;
  const fullAPI = true;

  const initialDoc = "docs/test.docx";

  const path = '/webviewer/';

  useEffect(() => {
    if (!hasBeenInitialized.current) {
      hasBeenInitialized.current = true;

      WebViewer.default(
        {
          path,
          licenseKey,
          initialDoc,
          fullAPI,
          enableOfficeEditing: true,
          officeMode: 'docx'
        },
        viewer.current
           ).then(async (instance) => {
        const { dispose } = attachDirtyTracker(instance, {
          onDirty: (source) => {
            console.log(`[DIRTY] flag set by: ${source}`);
            setDirty(true);
          },
          onSignal: (source, detail) => {
            console.log(`[signal] ${source}`, detail ?? '');
          },
        });
        return () => dispose();
      });
    }
  }, []);
  return (
    <div style={{ width: '100vw', height: '100vh', position: 'relative' }}>
      <div style={{
        position: 'absolute',
        top: 10,
        right: 10,
        zIndex: 999,
        padding: '8px 16px',
        borderRadius: 4,
        background: dirty ? '#e53e3e' : '#38a169',
        color: 'white',
        fontWeight: 'bold',
        fontSize: 14,
      }}>
        {dirty ? 'DIRTY' : 'CLEAN'}
      </div>
      <div style={{ width: '100%', height: '100%' }} ref={viewer}></div>
    </div>
  );
}
export default App