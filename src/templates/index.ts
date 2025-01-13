import pptxgen from 'pptxgenjs';

// Import the types from pptxgenjs
type BackgroundProps = pptxgen.BackgroundProps;
type ShapeFillProps = pptxgen.ShapeFillProps;
type ShapeProps = pptxgen.ShapeProps;

// Helper functions
function cleanText(text: string | undefined): string {
  if (!text) return '';
  return text.replace(/[^\w\s.,?!-]/g, '').trim();
}

// Update the background setting in addBackground function
function addBackground(slide: pptxgen.Slide, theme?: 'light' | 'dark' | 'gradient') {
  if (theme === 'dark') {
    slide.background = { color: '1a1a1a' } as BackgroundProps;
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
    } as BackgroundProps;
  } else {
    slide.background = { color: 'FFFFFF' } as BackgroundProps;
  }
}

// Update the shape fill properties
const gradientFill: ShapeFillProps = {
  type: 'solid',
  color: '2a2a40'
};

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
    color?: string;
  };
  subtitle?: string;
}

// Constants for better layout management
const SLIDE_MARGIN = 0.5;
const TITLE_HEIGHT = 0.8;
const CONTENT_START = SLIDE_MARGIN + TITLE_HEIGHT;
const SLIDE_WIDTH = 10 - (2 * SLIDE_MARGIN);

// Add this helper function after cleanText
function addTitle(slide: pptxgen.Slide, title: string | undefined, theme?: 'light' | 'dark' | 'gradient') {
  if (title) {
    slide.addText(cleanText(title), {
      x: SLIDE_MARGIN,
      y: SLIDE_MARGIN,
      w: SLIDE_WIDTH,
      h: TITLE_HEIGHT,
      fontSize: 32,
      bold: true,
      color: theme === 'dark' || theme === 'gradient' ? 'FFFFFF' : '000000'
    });
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
        fill: gradientFill
      } as ShapeProps);
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
        fill: gradientFill
      } as ShapeProps);
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
  },

  twoColumns: async (pres: pptxgen, data: SlideData) => {
    const slide = pres.addSlide();
    addBackground(slide, data.theme);
    addTitle(slide, data.title, data.theme);

    const points = data.content.section[0].bulletPoints?.[0]?.point;
    if (points) {
      const columnWidth = SLIDE_WIDTH / 2 - 0.1;
      points.forEach((point, idx) => {
        const col = idx % 2;
        const row = Math.floor(idx / 2);
        slide.addText(cleanText(point), {
          x: SLIDE_MARGIN + col * (columnWidth + 0.1),
          y: CONTENT_START + row * 0.8,
          w: columnWidth,
          h: 0.6,
          fontSize: 14,
          bullet: true,
          color: data.theme === 'dark' || data.theme === 'gradient' ? 'CCCCCC' : '666666'
        });
      });
    }
  },

  threeColumns: async (pres: pptxgen, data: SlideData) => {
    const slide = pres.addSlide();
    addBackground(slide, data.theme);
    addTitle(slide, data.title, data.theme);

    const points = data.content.section[0].bulletPoints?.[0]?.point;
    if (points) {
      const columnWidth = SLIDE_WIDTH / 3 - 0.1;
      points.forEach((point, idx) => {
        const col = idx % 3;
        const row = Math.floor(idx / 3);
        slide.addText(cleanText(point), {
          x: SLIDE_MARGIN + col * (columnWidth + 0.1),
          y: CONTENT_START + row * 0.8,
          w: columnWidth,
          h: 0.6,
          fontSize: 14,
          bullet: true,
          color: data.theme === 'dark' || data.theme === 'gradient' ? 'CCCCCC' : '666666'
        });
      });
    }
  },

  verticalTimeline: async (pres: pptxgen, data: SlideData) => {
    const slide = pres.addSlide();
    addBackground(slide, data.theme);
    addTitle(slide, data.title, data.theme);

    const points = data.content.section[0].bulletPoints?.[0]?.point;
    if (points) {
      const timelineX = SLIDE_MARGIN + 0.5;
      points.forEach((point, idx) => {
        // Add timeline point
        slide.addShape(pres.ShapeType.rect, {
          x: timelineX,
          y: CONTENT_START + idx * 1.2,
          w: 0.2,
          h: 0.2,
          fill: { type: 'solid', color: data.theme === 'dark' ? 'FFFFFF' : '000000' },
          roundLength: 0.2
        } as ShapeProps);

        // Add text
        slide.addText(cleanText(point), {
          x: timelineX + 0.4,
          y: CONTENT_START + idx * 1.2 - 0.1,
          w: SLIDE_WIDTH - 1,
          h: 0.6,
          fontSize: 14,
          color: data.theme === 'dark' || data.theme === 'gradient' ? 'CCCCCC' : '666666'
        });
      });
    }
  },

  horizontalTimeline: async (pres: pptxgen, data: SlideData) => {
    const slide = pres.addSlide();
    addBackground(slide, data.theme);
    addTitle(slide, data.title, data.theme);

    const points = data.content.section[0].bulletPoints?.[0]?.point;
    if (points) {
      const timelineY = CONTENT_START + 0.5;
      const itemWidth = SLIDE_WIDTH / points.length;
      
      points.forEach((point, idx) => {
        // Add timeline point
        slide.addShape(pres.ShapeType.rect, {
          x: SLIDE_MARGIN + idx * itemWidth + itemWidth/2 - 0.1,
          y: timelineY,
          w: 0.2,
          h: 0.2,
          fill: { type: 'solid', color: data.theme === 'dark' ? 'FFFFFF' : '000000' },
          roundLength: 0.2
        } as ShapeProps);

        // Add text
        slide.addText(cleanText(point), {
          x: SLIDE_MARGIN + idx * itemWidth,
          y: timelineY + 0.4,
          w: itemWidth,
          h: 0.8,
          fontSize: 12,
          align: 'center',
          color: data.theme === 'dark' || data.theme === 'gradient' ? 'CCCCCC' : '666666'
        });
      });
    }
  },
};