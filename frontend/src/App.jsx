import { useState } from 'react'
import Landing from './Landing'
import Tool from './Tool'

export default function App() {
  const [page, setPage] = useState('home')
  if (page === 'tool') return <Tool onBack={() => setPage('home')} />
  return <Landing onLaunch={() => setPage('tool')} />
}
