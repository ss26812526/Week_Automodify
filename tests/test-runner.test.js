(function () {
  var isNode = typeof module === 'object' && module.exports;
  var T = isNode ? require('../src/test-runner.js') : self.TestRunner;
  var describe = T.describe, it = T.it;
  var assertEqual = T.assertEqual, assertDeepEqual = T.assertDeepEqual, assertThrows = T.assertThrows;

  describe('測試框架本身', function () {
    it('assertEqual 相等時不擲出例外', function () {
      assertEqual(1, 1);
      assertEqual('一', '一');
    });

    it('assertEqual 不相等時會擲出例外', function () {
      assertThrows(function () { assertEqual(1, 2); });
    });

    it('assertDeepEqual 可比較陣列內容', function () {
      assertDeepEqual([1, 2, 3], [1, 2, 3]);
      assertThrows(function () { assertDeepEqual([1, 2], [1, 3]); });
    });

    it('支援 async 測試', async function () {
      var value = await Promise.resolve(42);
      assertEqual(value, 42);
    });
  });
})();
