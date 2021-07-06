import * as strings from 'SCWWebPartStrings';
import * as english from '../loc/en-us.js';
import * as french from '../loc/fr-fr.js';

export function SelectLanguage(lang) {
  switch(lang) {
    case "en-us": {
      return english;
    }
    case "fr-fr": {
      return french;
    }
    default: {
      return strings;
    }
 }
}
