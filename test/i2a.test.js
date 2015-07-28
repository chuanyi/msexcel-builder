var assert = require('assert');

tool = {
  i2a: function(i) {
    var alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ',
        len = alphabet.length;


        function getCellAdr(num){
          var pos = num % len,
              tmp = Math.floor(num/len),
              pos = (pos == 0) ? len : pos,
              tmp = (tmp > 0 && num % len==0) ? tmp - 1: tmp,
              output = alphabet.charAt(pos - 1);

          if(tmp > 0){
            output = getCellAdr(tmp) + output;
          };
          return output;
        }
    return getCellAdr(i);
  }
}

describe("i2a", function(){
  it ("matches and non-matches appropriately", function() {
    assert( tool.i2a(25) == "Y", tool.i2a(25) +" not equial to 'Y'"+ "  "+ 25);
    assert( tool.i2a(26) == "Z", tool.i2a(26) +" not equial to 'Z'"+ "  "+ 26);
    assert( tool.i2a(27) == "AA", tool.i2a(27) +" not equial to 'AA'"+ "  "+ 27);
    assert( tool.i2a(52) == "AZ", tool.i2a(52) +" not equial to 'AZ'"+ "  "+ 52);
    assert( tool.i2a(53) == "BA", tool.i2a(53) +" not equial to 'BA'"+ "  "+ 53);
    assert( tool.i2a(78) == "BZ", tool.i2a(78) +" not equial to 'BZ'"+ "  "+ 78);
    assert( tool.i2a(79) == "CA", tool.i2a(79) +" not equial to 'CA'"+ "  "+ 79);
    assert( tool.i2a(702) == "ZZ", tool.i2a(702) +" not equial to 'ZZ'"+ "  "+ 702);
    assert( tool.i2a(703) == "AAA", tool.i2a(703) +" not equial to 'AAA'"+ "  "+ 703);
  });
});

