import { Link } from 'react-router-dom';
import { ArrowRight, FileSpreadsheet, Image, Presentation } from 'lucide-react';

export default function Home() {
  return (
    <div className="flex flex-col items-center justify-center min-h-[calc(100vh-200px)] px-4">
      <div className="text-center max-w-2xl">
        <div className="flex justify-center gap-4 mb-8">
          <FileSpreadsheet className="h-12 w-12 text-primary" />
          <Image className="h-12 w-12 text-primary" />
          <Presentation className="h-12 w-12 text-primary" />
        </div>

        <h1 className="text-4xl font-bold tracking-tight mb-4">
          HR PowerPoint Generator
        </h1>

        <p className="text-lg text-muted-foreground mb-8">
          Generate professional employee PowerPoint presentations directly from
          Excel and ZIP photo files.
        </p>

        <Link
          to="/generate"
          className="inline-flex items-center gap-2 bg-primary text-primary-foreground px-6 py-3 rounded-lg font-medium hover:bg-primary/90 transition-colors"
        >
          Start
          <ArrowRight className="h-4 w-4" />
        </Link>
      </div>
    </div>
  );
}
