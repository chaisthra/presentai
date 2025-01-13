import { useState } from 'react';
import { generateOutline, generateSlideContent, formatCustomContent } from '../lib/gemini';
import pptxgen from 'pptxgenjs';
import { templates } from '../templates';
import toast from 'react-hot-toast';
import { twMerge } from 'tailwind-merge';

const TEMPLATE_TYPES = [
  { id: 'comprehensive', name: 'Comprehensive Layout' },
  { id: 'twoColumns', name: 'Two Columns' },
  { id: 'threeColumns', name: 'Three Columns' },
  { id: 'verticalTimeline', name: 'Vertical Timeline' },
  { id: 'horizontalTimeline', name: 'Horizontal Timeline' }
] as const;

const THEMES = [
  { id: 'light', name: 'Light Theme' },
  { id: 'dark', name: 'Dark Theme' },
  { id: 'gradient', name: 'Gradient Theme' }
] as const;

export function PresentationGenerator() {
  const [topic, setTopic] = useState('');
  const [numSlides, setNumSlides] = useState(5);
  const [selectedTheme, setSelectedTheme] = useState<'light' | 'dark' | 'gradient'>('light');
  const [outline, setOutline] = useState<string[]>([]);
  const [customOutline, setCustomOutline] = useState('');
  const [customContent, setCustomContent] = useState('');
  const [selectedTemplates, setSelectedTemplates] = useState<Record<number, string>>({});
  const [isGenerating, setIsGenerating] = useState(false);
  const [includeTitleSlide, setIncludeTitleSlide] = useState(true);
  const [includeTableOfContents, setIncludeTableOfContents] = useState(true);
  const [includeThankYou, setIncludeThankYou] = useState(true);
  const [useCustomContent, setUseCustomContent] = useState(false);

  const handleGenerateOutline = async () => {
    if (!topic && !customOutline) {
      toast.error('Please enter a topic or custom outline');
      return;
    }

    try {
      setIsGenerating(true);
      let titles: string[];
      
      if (customOutline) {
        titles = customOutline.split('\n').map(line => line.trim()).filter(line => line.length > 0);
      } else {
        titles = await generateOutline(topic, numSlides);
      }
      
      setOutline(titles);
      setNumSlides(titles.length);

      const defaults = titles.reduce((acc, _, index) => {
        acc[index] = 'comprehensive';
        return acc;
      }, {} as Record<number, string>);
      setSelectedTemplates(defaults);

      toast.success('Outline generated successfully');
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Failed to generate outline');
    } finally {
      setIsGenerating(false);
    }
  };

  const handleGeneratePresentation = async () => {
    if (!outline.length) {
      toast.error('Please generate an outline first');
      return;
    }

    try {
      setIsGenerating(true);
      const pres = new pptxgen();

      // Add title slide if enabled
      if (includeTitleSlide) {
        await templates.titleSlide(pres, {
          type: 'titleSlide',
          title: topic,
          content: { section: [{ paragraph: [''] }] },
          theme: selectedTheme,
          titleOptions: {
            fontSize: 48,
            bold: true,
            color: selectedTheme === 'dark' ? 'FFFFFF' : '000000'
          }
        });
      }

      // Add table of contents if enabled
      if (includeTableOfContents) {
        await templates.comprehensive(pres, {
          type: 'comprehensive',
          title: 'Table of Contents',
          content: {
            section: [{
              bulletPoints: [{
                point: outline,
              }],
            }],
          },
          theme: selectedTheme,
        });
      }

      // Generate content slides
      for (let i = 0; i < outline.length; i++) {
        const title = outline[i];
        const templateType = selectedTemplates[i] || 'comprehensive';
        const template = templates[templateType as keyof typeof templates];

        if (!template) {
          throw new Error(`Template ${templateType} not found`);
        }

        let content;
        if (useCustomContent && customContent) {
          content = await formatCustomContent(customContent, title);
        } else {
          content = await generateSlideContent(title, topic);
        }

        await template(pres, {
          type: templateType,
          title,
          content: {
            section: [{
              paragraph: [content.intro],
              bulletPoints: [{
                point: content.points,
              }],
            }],
          },
          theme: selectedTheme,
        });
      }

      // Add thank you slide if enabled
      if (includeThankYou) {
        await templates.thankYou(pres, {
          type: 'thankYou',
          content: { section: [{ paragraph: [''] }] },
          theme: selectedTheme,
          titleOptions: {
            fontSize: 48,
            bold: true,
            color: selectedTheme === 'dark' ? 'FFFFFF' : '000000'
          }
        });
      }

      const fileName = `${topic.replace(/[^a-z0-9]/gi, '_').toLowerCase()}_presentation.pptx`;
      await pres.writeFile({ fileName });
      toast.success('Presentation generated successfully!');
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Failed to generate presentation');
    } finally {
      setIsGenerating(false);
    }
  };

  return (
    <div className="glass-card rounded-2xl p-8">
      {/* Topic Input */}
      <div className="mb-6">
        <label className="block text-sm font-medium text-white/80 mb-2">
          Presentation Topic
        </label>
        <input
          type="text"
          value={topic}
          onChange={(e) => setTopic(e.target.value)}
          className="w-full rounded-lg glass-input"
          placeholder="Enter your presentation topic"
          disabled={useCustomContent}
        />
      </div>

      {/* Custom Content Toggle */}
      <div className="mb-6">
        <label className="flex items-center">
          <input
            type="checkbox"
            checked={useCustomContent}
            onChange={(e) => setUseCustomContent(e.target.checked)}
            className="mr-2"
          />
          <span className="text-sm text-white/80">Use Custom Content</span>
        </label>
      </div>

      {useCustomContent ? (
        <>
          {/* Custom Outline */}
          <div className="mb-6">
            <label className="block text-sm font-medium text-white/80 mb-2">
              Custom Outline (One title per line)
            </label>
            <textarea
              value={customOutline}
              onChange={(e) => setCustomOutline(e.target.value)}
              className="w-full h-32 rounded-lg glass-input"
              placeholder="Enter your custom outline..."
            />
          </div>

          {/* Custom Content */}
          <div className="mb-6">
            <label className="block text-sm font-medium text-white/80 mb-2">
              Custom Content Template
            </label>
            <textarea
              value={customContent}
              onChange={(e) => setCustomContent(e.target.value)}
              className="w-full h-32 rounded-lg glass-input"
              placeholder="Enter your custom content template..."
            />
          </div>
        </>
      ) : (
        <div className="mb-6">
          <label className="block text-sm font-medium text-white/80 mb-2">
            Number of Slides
          </label>
          <input
            type="number"
            min="1"
            max="20"
            value={numSlides}
            onChange={(e) => setNumSlides(Math.max(1, Math.min(20, parseInt(e.target.value) || 5)))}
            className="w-full rounded-lg glass-input"
          />
        </div>
      )}

      {/* Theme Selection */}
      <div className="mb-6">
        <label className="block text-sm font-medium text-white/80 mb-2">
          Presentation Theme
        </label>
        <div className="grid grid-cols-3 gap-4">
          {THEMES.map((theme) => (
            <button
              key={theme.id}
              onClick={() => setSelectedTheme(theme.id as 'light' | 'dark' | 'gradient')}
              className={twMerge(
                'py-2 px-4 rounded-lg glass-button',
                selectedTheme === theme.id && 'bg-blue-500/30 border-blue-400/50'
              )}
            >
              {theme.name}
            </button>
          ))}
        </div>
      </div>

      {/* Optional Slides */}
      <div className="mb-6">
        <label className="block text-sm font-medium text-white/80 mb-2">
          Optional Slides
        </label>
        <div className="flex gap-4">
          <label className="flex items-center">
            <input
              type="checkbox"
              checked={includeTitleSlide}
              onChange={(e) => setIncludeTitleSlide(e.target.checked)}
              className="mr-2"
            />
            Title Slide
          </label>
          <label className="flex items-center">
            <input
              type="checkbox"
              checked={includeTableOfContents}
              onChange={(e) => setIncludeTableOfContents(e.target.checked)}
              className="mr-2"
            />
            Table of Contents
          </label>
          <label className="flex items-center">
            <input
              type="checkbox"
              checked={includeThankYou}
              onChange={(e) => setIncludeThankYou(e.target.checked)}
              className="mr-2"
            />
            Thank You Slide
          </label>
        </div>
      </div>

      {outline.length > 0 && (
        <div className="mb-6">
          <label className="block text-sm font-medium text-white/80 mb-2">
            Slide Templates
          </label>
          <div className="space-y-4">
            {outline.map((title, index) => (
              <div key={index} className="flex gap-4 items-center">
                <span className="text-sm text-white/60 w-8">{index + 1}.</span>
                <div className="flex-1 flex gap-4">
                  <select
                    value={selectedTemplates[index] || 'comprehensive'}
                    onChange={(e) => setSelectedTemplates(prev => ({
                      ...prev,
                      [index]: e.target.value
                    }))}
                    className="w-64 rounded-lg glass-input"
                  >
                    {TEMPLATE_TYPES.map(template => (
                      <option key={template.id} value={template.id}>
                        {template.name}
                      </option>
                    ))}
                  </select>
                  <span className="flex-1 text-white/80 text-sm py-2">{title}</span>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Action Buttons */}
      <div className="flex gap-4">
        <button
          onClick={handleGenerateOutline}
          disabled={isGenerating || (!topic && !customOutline)}
          className="flex-1 py-3 px-6 rounded-lg glass-button"
        >
          {isGenerating ? 'Generating Outline...' : 'Generate Outline'}
        </button>
        <button
          onClick={handleGeneratePresentation}
          disabled={isGenerating || !outline.length}
          className="flex-1 py-3 px-6 rounded-lg glass-button glow"
        >
          {isGenerating ? 'Generating Presentation...' : 'Generate Presentation'}
        </button>
      </div>
    </div>
  );
}