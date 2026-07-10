import { Routes, Route } from 'react-router-dom';
import Layout from './components/Layout';
import Home from './pages/Home';
import GeneratePPT from './pages/GeneratePPT';

function App() {
  return (
    <Layout>
      <Routes>
        <Route path="/" element={<Home />} />
        <Route path="/generate" element={<GeneratePPT />} />
      </Routes>
    </Layout>
  );
}

export default App;
