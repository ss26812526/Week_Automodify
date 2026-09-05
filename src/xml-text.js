// XML 文字的跳脫與還原
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.XmlText = factory();
})(typeof self !== 'undefined' ? self : this, function () {

  function escape(s) {
    if (s === null || s === undefined) return '';
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function unescape(s) {
    if (s === null || s === undefined) return '';
    return String(s)
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&#x([0-9a-fA-F]+);/g, function (_, hex) {
        return String.fromCodePoint(parseInt(hex, 16));
      })
      .replace(/&#([0-9]+);/g, function (_, dec) {
        return String.fromCodePoint(parseInt(dec, 10));
      })
      .replace(/&amp;/g, '&');
  }

  return { escape: escape, unescape: unescape };
});
