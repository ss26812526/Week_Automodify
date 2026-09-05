(function () {
  var isNode = typeof module === 'object' && module.exports;
  var T = isNode ? require('../src/test-runner.js') : self.TestRunner;
  var X = isNode ? require('../src/xml-text.js') : self.XmlText;
  var describe = T.describe, it = T.it;
  var assertEqual = T.assertEqual;

  describe('XML 文字處理', function () {
    it('還原五種實體', function () {
      assertEqual(X.unescape('a&amp;b&lt;c&gt;d&quot;e&apos;f'), 'a&b<c>d"e\'f');
    });

    it('跳脫五種字元', function () {
      assertEqual(X.escape('a&b<c>d"e\'f'), 'a&amp;b&lt;c&gt;d&quot;e&apos;f');
    });

    it('還原數值實體', function () {
      assertEqual(X.unescape('&#65;&#x42;'), 'AB');
    });

    it('中文不受影響', function () {
      assertEqual(X.escape('機房檢查'), '機房檢查');
      assertEqual(X.unescape('機房檢查'), '機房檢查');
    });

    it('跳脫後再還原可回到原文', function () {
      var s = '備註 <重要> & "緊急"';
      assertEqual(X.unescape(X.escape(s)), s);
    });

    it('先還原 amp 會出錯，順序必須正確', function () {
      assertEqual(X.unescape('&amp;lt;'), '&lt;');
    });

    it('空字串與 null 安全', function () {
      assertEqual(X.escape(''), '');
      assertEqual(X.unescape(''), '');
      assertEqual(X.escape(null), '');
      assertEqual(X.unescape(null), '');
    });
  });
})();
