// https://github.com/artofthesmart/QUnitGS2
// https://qunitjs.com/

var QUnit = QUnitGS2.QUnit;

function doGet() {

  QUnitGS2.init();
  QUnit.module("Basic tests");

  QUnit.test("Get properties", function (assert) {
    cofg = getProperties_();
    assert.equal(cfg.onenoteMail, 'me@onenote.com')
  });

  QUnit.test("Compose header", function (assert) {
    // Arrange
    const lastMsgMock = {
      getFrom: function() { return "m@v.com"; },
      getTo: function() { return "v@m.com"; }
    };
    const hdrFields = ['from', 'to'];

    // Act
    let header = composeHeader_(hdrFields, lastMsgMock);

    // Assert
    assert.equal(header, "From: m@v.com\nTo: v@m.com\n");
  });

  QUnit.start();
  return QUnitGS2.getHtml();
}

function getResultsFromServer() {
  return QUnitGS2.getResultsFromServer();
}