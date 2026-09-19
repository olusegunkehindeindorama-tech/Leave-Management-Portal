function getIndexHtml_() {
  var parts = []
    .concat(indexHtmlPart0_())
    .concat(indexHtmlPart1_())
    .concat(indexHtmlPart2_())
    .concat(indexHtmlPart3_());
  return Utilities.newBlob(Utilities.base64Decode(parts.join(''))).getDataAsString('UTF-8');
}
