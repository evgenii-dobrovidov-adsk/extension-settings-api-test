import { useEffect, useState } from 'react'
import { CssBaseline } from '@weave-mui/material'
import { ThemeProvider, createTheme, getTheme, densities, themes } from '@weave-mui/styles'
import App from './App.tsx'
import './typography.css'

function WeaveRoot() {
  const [theme, setTheme] = useState(() => createTheme({}))

  useEffect(() => {
    let cancelled = false

    async function loadTheme() {
      try {
        const baseTheme = createTheme({})
        const weaveTheme = await getTheme(themes.LIGHT_GRAY, densities.HIGH)

        if (!cancelled) {
          setTheme(createTheme(baseTheme, weaveTheme))
        }
      } catch {
        if (!cancelled) {
          setTheme(createTheme({}))
        }
      }
    }

    void loadTheme()

    return () => {
      cancelled = true
    }
  }, [])

  return (
    <ThemeProvider theme={theme}>
      <CssBaseline />
      <App />
    </ThemeProvider>
  )
}

export default WeaveRoot
