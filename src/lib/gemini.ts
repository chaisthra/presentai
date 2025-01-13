import { GoogleGenerativeAI } from '@google/generative-ai';

const API_KEY = import.meta.env.VITE_GEMINI_API_KEY;

if (!API_KEY) {
  throw new Error('VITE_GEMINI_API_KEY is not defined in environment variables');
}

const genAI = new GoogleGenerativeAI(API_KEY);

export async function generateOutline(topic: string, numSlides: number = 5): Promise<string[]> {
  try {
    const model = genAI.getGenerativeModel({ model: 'gemini-pro' });

    const prompt = `Create a presentation outline for the topic: "${topic}". 
      Generate exactly ${numSlides} slide titles that would make a compelling presentation. 
      Keep each title under 50 characters.
      Return only the slide titles, one per line.`;

    const result = await model.generateContent(prompt);
    const response = await result.response;
    const titles = response.text()
      .split('\n')
      .map(line => line.trim())
      .filter(line => line.length > 0)
      .map(line => line.replace(/[^\w\s-]/g, ''));
    
    if (titles.length !== numSlides) {
      throw new Error('Generated outline does not match requested number of slides');
    }
    
    return titles;
  } catch (error) {
    console.error('Error generating outline:', error);
    throw new Error('Failed to generate outline. Please try again.');
  }
}

interface SlideContent {
  intro: string;
  points: string[];
  visualSuggestion: string;
}

export async function generateSlideContent(title: string, topic: string): Promise<SlideContent> {
  try {
    const model = genAI.getGenerativeModel({ model: 'gemini-pro' });

    const prompt = `Create concise content for a presentation slide titled "${title}" about ${topic}.
      Rules:
      - Keep the intro under 150 characters
      - Each bullet point must be under 100 characters
      - Use simple language
      - No special characters or symbols
      - No quotation marks
      
      Structure your response exactly like this:
      {
        "intro": "Brief clear introduction",
        "points": [
          "First key point",
          "Second key point",
          "Third key point"
        ],
        "visualSuggestion": "A simple visual suggestion"
      }`;

    const result = await model.generateContent(prompt);
    const response = await result.response;
    const text = response.text().trim();
    
    try {
      const content = JSON.parse(text);
      
      if (!content.intro || !Array.isArray(content.points) || !content.visualSuggestion) {
        throw new Error('Invalid response structure');
      }

      // Clean and format the content
      const cleanedContent = {
        intro: content.intro
          .replace(/[^\w\s.,]/g, '')
          .trim()
          .slice(0, 150),
        points: content.points
          .slice(0, 3)
          .map(point => point
            .replace(/[^\w\s.,]/g, '')
            .trim()
            .slice(0, 100)
          ),
        visualSuggestion: content.visualSuggestion
          .replace(/[^\w\s.,]/g, '')
          .trim()
      };

      // Ensure exactly 3 points
      while (cleanedContent.points.length < 3) {
        cleanedContent.points.push("Additional point needed");
      }

      return cleanedContent;
    } catch (e) {
      console.error('JSON parsing error:', e);
      return {
        intro: "Content generation failed. Please try again.",
        points: [
          "Point 1",
          "Point 2",
          "Point 3"
        ],
        visualSuggestion: "Generic visual element"
      };
    }
  } catch (error) {
    console.error('Error generating slide content:', error);
    throw new Error('Failed to generate slide content. Please try again.');
  }
}

export async function formatCustomContent(content: string, title: string): Promise<SlideContent> {
  try {
    const model = genAI.getGenerativeModel({ model: 'gemini-pro' });

    const prompt = `Format this content for a presentation slide titled "${title}".
      Rules:
      - Keep the intro under 150 characters
      - Each bullet point must be under 100 characters
      - Use simple language
      - No special characters or symbols
      - No quotation marks

      Content to format:
      ${content}
      
      Structure your response exactly like this:
      {
        "intro": "Brief clear introduction",
        "points": [
          "First key point",
          "Second key point",
          "Third key point"
        ],
        "visualSuggestion": "A simple visual suggestion"
      }`;

    const result = await model.generateContent(prompt);
    const response = await result.response;
    const text = response.text().trim();
    
    try {
      const formattedContent = JSON.parse(text);
      
      if (!formattedContent.intro || !Array.isArray(formattedContent.points) || !formattedContent.visualSuggestion) {
        throw new Error('Invalid response structure');
      }

      // Clean and format the content
      const cleanedContent = {
        intro: formattedContent.intro
          .replace(/[^\w\s.,]/g, '')
          .trim()
          .slice(0, 150),
        points: formattedContent.points
          .slice(0, 3)
          .map(point => point
            .replace(/[^\w\s.,]/g, '')
            .trim()
            .slice(0, 100)
          ),
        visualSuggestion: formattedContent.visualSuggestion
          .replace(/[^\w\s.,]/g, '')
          .trim()
      };

      // Ensure exactly 3 points
      while (cleanedContent.points.length < 3) {
        cleanedContent.points.push("Additional point needed");
      }

      return cleanedContent;
    } catch (e) {
      console.error('JSON parsing error:', e);
      return {
        intro: "Content formatting failed. Please try again.",
        points: [
          "Point 1",
          "Point 2",
          "Point 3"
        ],
        visualSuggestion: "Generic visual element"
      };
    }
  } catch (error) {
    console.error('Error formatting content:', error);
    throw new Error('Failed to format content. Please try again.');
  }
}