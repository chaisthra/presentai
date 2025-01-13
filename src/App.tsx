import { Navbar } from './components/Navbar';
import { PresentationGenerator } from './components/PresentationGenerator';

function App() {
  return (
    <>
      <video
        autoPlay
        loop
        muted
        playsInline
        className="video-background"
      >
        <source
          src="https://videos.pexels.com/video-files/18069232/18069232-uhd_2560_1440_24fps.mp4"
          type="video/mp4"
        />
      </video>

      <Navbar />

      <div className="min-h-screen pt-24 pb-12 px-4 sm:px-6 lg:px-8 relative">
        <div className="max-w-5xl mx-auto">
          <div className="text-center mb-12">
            <h1 className="text-5xl font-bold heading-gradient mb-4">Create Stunning Presentations</h1>
            <p className="text-lg text-white/80">Transform your ideas into professional presentations with AI assistance</p>
          </div>

          <PresentationGenerator />
        </div>
      </div>
    </>
  );
}

export default App;