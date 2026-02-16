import fs from 'fs';
import path from 'path';

const urlFilePath = path.join(__dirname, '../url.txt');
export const BASE_URL = fs.readFileSync(urlFilePath, 'utf8').trim();
