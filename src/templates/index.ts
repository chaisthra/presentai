import pptxgen from 'pptxgenjs';

export interface SlideData {
  type: string;
  title?: string;
  content: {
    section: Array<{
      heading?: string;
      paragraph?: string[];
      bulletPoints?: Array<{
        point: string[];
      }>;
    }>;
  };
  theme?: 'light' | 'dark' | 'gradient';
  titleOptions?: {
    fontSize?: number;
    bold?: boolean;
    alignment?: 'left' | 'center' | 'right';
  };
  subtitle?: string;
}

// Constants for better layout management
const SLIDE_MARGIN = 0.5;
const TITLE_HEIGHT = 0.8;
const CONTENT_START = SLIDE_MARGIN + TITLE_HEIGHT;
const SLIDE_WIDTH = 10 - (2 * SLIDE_MARGIN);
const SLIDE_HEIGHT = 7.5 - (2 * SLIDE_MARGIN);

// Helper functions
function cleanText(text: string | undefined): string {
  if (!text) return '';
  return text.replace(/[^\w\s.,?!-]/g, '').trim();
}

function addBackground(slide: pptxgen.Slide, theme?: 'light' | 'dark' | 'gradient') {
  if (theme === 'dark') {
    slide.background = { color: '1a1a1a' };
  } else if (theme === 'gradient') {
    slide.background = { 
      color: '1a1a1a',
      gradient: {
        type: 'linear',
        stops: [
          { color: '1a1a1a', position: 0 },
          { color: '2a2a40', position: 1 }
        ]
      }
    };
  } else {
    slide.background = { color: 'FFFFFF' };
  }
}

export const templates = {
  titleSlide: async (pres: pptxgen, data: SlideData) => {
    const slide = pres.addSlide();
    addBackground(slide, data.theme);
    
    // Add decorative background elements
    if (data.theme === 'gradient') {
      slide.addShape(pres.ShapeType.rect, {
        x: 0,
        y: 0,
        w: '100%',
        h: '100%',
        fill: {
          type: 'gradient',
          color: ['2a2a40', '1a1a2e'],
          angle: 45
        },
        opacity: 0.7
      });
    }

    // Add main title with enhanced styling
    slide.addText(cleanText(data.title), {
      x: '10%',
      y: '40%',
      w: '80%',
      h: '20%',
      fontSize: data.titleOptions?.fontSize || 54,
      bold: data.titleOptions?.bold ?? true,
      align: data.titleOptions?.alignment || 'center',
      color: data.theme === 'dark' || data.theme === 'gradient' ? 'FFFFFF' : '000000',
      shadow: { type: 'outer', blur: 3, offset: 2, angle: 45, color: '808080', opacity: 0.3 }
    });

    // Add subtitle if provided
    if (data.subtitle) {
      slide.addText(cleanText(data.subtitle), {
        x: '10%',
        y: '60%',
        w: '80%',
        h: '10%',
        fontSize: 28,
        color: data.theme === 'dark' || data.theme === 'gradient' ? 'CCCCCC' : '666666',
        align: 'center'
      });
    }
  },

  thankYou: async (pres: pptxgen, data: SlideData) => {
    const slide = pres.addSlide();
    addBackground(slide, data.theme);
    
    // Add decorative background
    if (data.theme === 'gradient') {
      slide.addShape(pres.ShapeType.rect, {
        x: 0,
        y: 0,
        w: '100%',
        h: '100%',
        fill: {
          type: 'gradient',
          color: ['2a2a40', '1a1a2e'],
          angle: 45
        },
        opacity: 0.7
      });
    }

    // Add "Thank You" text with enhanced styling
    slide.addText('Thank You!', {
      x: '10%',
      y: '40%',
      w: '80%',
      h: '20%',
      fontSize: 54,
      bold: true,
      align: 'center',
      color: data.theme === 'dark' || data.theme === 'gradient' ? 'FFFFFF' : '000000',
      shadow: { type: 'outer', blur: 3, offset: 2, angle: 45, color: '808080', opacity: 0.3 }
    });

    // Add subtitle
    slide.addText('Questions & Discussion', {
      x: '10%',
      y: '60%',
      w: '80%',
      h: '10%',
      fontSize: 32,
      color: data.theme === 'dark' || data.theme === 'gradient' ? 'CCCCCC' : '666666',
      align: 'center'
    });
  },

  comprehensive: async (pres: pptxgen, data: SlideData) => {
    const slide = pres.addSlide();
    addBackground(slide, data.theme);
    
    // Add title
    if (data.title) {
      slide.addText(cleanText(data.title), {
        x: SLIDE_MARGIN,
        y: SLIDE_MARGIN,
        w: SLIDE_WIDTH,
        h: TITLE_HEIGHT,
        fontSize: 32,
        bold: true,
        color: data.theme === 'dark' || data.theme === 'gradient' ? 'FFFFFF' : '000000'
      });
    }
    
    // Add content
    const points = data.content.section[0].bulletPoints?.[0]?.point;
    if (points && Array.isArray(points)) {
      points.forEach((point, idx) => {
        slide.addText(cleanText(point), {
          x: SLIDE_MARGIN,
          y: CONTENT_START + (idx * 0.8),
          w: SLIDE_WIDTH,
          h: 0.6,
          fontSize: 18,
          bullet: true,
          color: data.theme === 'dark' || data.theme === 'gradient' ? 'CCCCCC' : '666666'
        });
      });
    }
  }
};