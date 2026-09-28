function getIndexHtml_() {
  var b64 = INDEX_B64_P0_ + INDEX_B64_P1_ + INDEX_B64_P2_ + INDEX_B64_P3_ + INDEX_B64_P4_;
  var bytes = Utilities.base64Decode(b64);
  return Utilities.newBlob(bytes).getDataAsString();
}
