import { parseString } from 'xml2js';

export const parseXMLFile = (xmlContent: string): Promise<any> => {
  return new Promise((resolve, reject) => {
    parseString(xmlContent, (err, result) => {
      if (err) {
        reject(err);
      } else {
        resolve(result);
      }
    });
  });
};