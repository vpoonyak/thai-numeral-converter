import { lazy, Suspense, useState } from 'react';
import { ThemeProvider, createTheme } from '@fluentui/react';
import TaskPane from './TaskPane';
import './App.css';

// The website view is only needed outside Word, so keep it out of the task pane bundle
const Landing = lazy(() => import('./Landing'));

const myTheme = createTheme({
  defaultFontStyle: { fontFamily: "'Segoe UI', 'Leelawadee UI', 'Sarabun', 'Thonburi', system-ui, sans-serif" },
  palette: {
    themePrimary: '#d83b01',
    themeLighterAlt: '#fdf6f2',
    themeLighter: '#f7dbcf',
    themeLight: '#f0beaa',
    themeTertiary: '#e18360',
    themeSecondary: '#dc5120',
    themeDarkAlt: '#c23501',
    themeDark: '#a42d01',
    themeDarker: '#792101',
    neutralLighterAlt: '#f8f8f8',
    neutralLighter: '#f4f4f4',
    neutralLight: '#eaeaea',
    neutralQuaternaryAlt: '#dadada',
    neutralQuaternary: '#d0d0d0',
    neutralTertiaryAlt: '#c8c8c8',
    neutralTertiary: '#a19f9d',
    neutralSecondary: '#605e5c',
    neutralPrimaryAlt: '#3b3a39',
    neutralPrimary: '#323130',
    neutralDark: '#201f1e',
    black: '#000000',
    white: '#ffffff',
  }
});

function App() {
  // Outside Word (e.g. the GitHub Pages site) Office.context.host is undefined
  const [isOffice] = useState(() => typeof Office !== 'undefined' && !!Office.context?.host);

  return (
    <ThemeProvider theme={myTheme}>
      {isOffice ? <TaskPane /> : <Suspense fallback={null}><Landing /></Suspense>}
    </ThemeProvider>
  );
}

export default App;
