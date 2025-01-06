import { useState } from 'react'
import './App.css'
import { BrowserRouter as Router, Routes, Route, Link } from 'react-router-dom'
import { PlaceholderComponent } from './components/placeholderComponent'

function App() {

  return (
    <>
      <Router>
        <Routes>
          <Route path="/:query" element={<PlaceholderComponent/>}></Route>
          <Route path="*" element={(
            <div>
              Error generating layout
            </div>
          )}/>
        </Routes>
      </Router>
    </>
  )
}

export default App
