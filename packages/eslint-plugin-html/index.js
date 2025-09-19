'use strict';

const SCRIPT_REGEX = /<script\b[^>]*>([\s\S]*?)(<\/script>|$)/gi;
const NEWLINE_REGEX = /\r\n|\r|\n/;
const offsets = new Map();

function computeLineAndColumn(text) {
  const lines = text.split(NEWLINE_REGEX);
  const line = lines.length;
  const column = lines[line - 1].length + 1;
  return { line, column };
}

module.exports = {
  processors: {
    html: {
      preprocess(text, filename) {
        const blocks = [];
        const metadata = [];
        let match;
        while ((match = SCRIPT_REGEX.exec(text))) {
          const scriptContent = match[1] || '';
          if (!scriptContent.trim()) {
            continue;
          }
          const startIndex = match.index + match[0].indexOf(scriptContent);
          const preceding = text.slice(0, startIndex);
          const position = computeLineAndColumn(preceding);
          metadata.push(position);
          blocks.push(scriptContent);
        }
        if (metadata.length) {
          offsets.set(filename, metadata);
          return blocks;
        }
        offsets.delete(filename);
        return [''];
      },
      postprocess(messageLists, filename) {
        const metadata = offsets.get(filename) || [];
        offsets.delete(filename);
        const results = [];
        messageLists.forEach((messages, index) => {
          const offset = metadata[index];
          messages.forEach((message) => {
            const adjusted = { ...message };
            if (offset) {
              if (typeof adjusted.line === 'number') {
                adjusted.line += offset.line - 1;
              }
              if (typeof adjusted.column === 'number') {
                const originalLine = message.line || 1;
                adjusted.column += originalLine === 1 ? offset.column - 1 : 0;
              }
              if (typeof adjusted.endLine === 'number') {
                adjusted.endLine += offset.line - 1;
              }
              if (typeof adjusted.endColumn === 'number') {
                const originalEndLine = message.endLine || message.line || 1;
                adjusted.endColumn += originalEndLine === 1 ? offset.column - 1 : 0;
              }
            }
            results.push(adjusted);
          });
        });
        return results;
      },
      supportsAutofix: false
    }
  }
};
