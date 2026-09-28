function getIndexHtml_() {
  var b64 = INDEX_B64_P0_ + INDEX_B64_P1_ + INDEX_B64_P2_ + INDEX_B64_P3_ + INDEX_B64_P4_ + INDEX_B64_P5_ + INDEX_B64_P6_ + INDEX_B64_P7_ + INDEX_B64_P8_;
  var bytes = Utilities.base64Decode(b64);
  return Utilities.newBlob(bytes).getDataAsString();
}
