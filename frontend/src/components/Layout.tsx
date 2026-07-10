import { Link, useLocation } from 'react-router-dom';
import { Presentation, Home } from 'lucide-react';
import { cn } from '../lib/utils';

interface LayoutProps {
  children: React.ReactNode;
}

export default function Layout({ children }: LayoutProps) {
  const location = useLocation();

  const navItems = [
    { path: '/', label: 'Home', icon: Home },
    { path: '/generate', label: 'Generate PPT', icon: Presentation },
  ];

  return (
    <div className="min-h-screen flex flex-col">
      <header className="border-b">
        <nav className="container mx-auto px-4 py-4 flex items-center justify-between">
          <Link to="/" className="flex items-center gap-2 font-semibold text-lg">
            <Presentation className="h-6 w-6 text-primary" />
            <span>HR PPT Generator</span>
          </Link>

          <div className="flex items-center gap-6">
            {navItems.map(({ path, label, icon: Icon }) => (
              <Link
                key={path}
                to={path}
                className={cn(
                  'flex items-center gap-2 text-sm font-medium transition-colors hover:text-primary',
                  location.pathname === path
                    ? 'text-primary'
                    : 'text-muted-foreground'
                )}
              >
                <Icon className="h-4 w-4" />
                {label}
              </Link>
            ))}
          </div>
        </nav>
      </header>

      <main className="flex-1">{children}</main>

      <footer className="border-t py-4 text-center text-sm text-muted-foreground">
        HR PowerPoint Generator
      </footer>
    </div>
  );
}
