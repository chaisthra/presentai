import React from 'react';
import logo from '/logo.svg';

export function Navbar() {
  return (
    <nav className="fixed top-0 left-0 right-0 z-50">
      <div className="glass-card border-b border-white/5">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex items-center h-16">
            <div className="flex items-center">
              <img src={logo} alt="PresentAI Logo" className="h-8 w-8" />
              <span className="ml-2 text-xl font-semibold text-white">PresentAI</span>
            </div>
          </div>
        </div>
      </div>
    </nav>
  );
}