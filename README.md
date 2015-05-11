== 引言 ==

转载于: https://gist.github.com/semimiracle/15b003c57abf11289def
最近要用Perl写个处理包含中文的XLSX文件的小脚本，折腾了好半天才搞定中文处理，记录下来方便日后查询。


== 背景 ==
* 读XLSX文件使用[[https://metacpan.org/pod/Spreadsheet::XLSX|Spreadsheet::XLSX]]
* * 写XLSX文件使用[[https://metacpan.org/pod/Excel::Writer::XLSX|Excel::Writer::XLSX]]
* *Microsoft Excel 2007* 之后的xlsx格式，实际是ZIP+XML（将xlsx文件重名名为zip，而后再用winzip提取文件就可看到内部的xml文件），所以只能是utf-8编码。

  --- eg.xml ---
  <hello>大家好<\hello>
  <world>我是周杰伦<\world>

  --- Binary Mode ---
  0000000: 3c68 656c 6c6f 3ee5 a4a7 e5ae b6e5 a5bd  <hello>.........
  0000010: 3c5c 6865 6c6c 6f3e 0d0a 3c77 6f72 6c64  <\hello>..<world
  0000020: 3ee6 8891 e698 afe5 91a8 e69d b0e4 bca6  >...............
  0000030: 3c5c 776f 726c 643e 0d0a                 <\world>..

  --- Utf-8 mapping table ---
  e5a4a7: 大
  e5aeb6: 家
  e5a5bd: 好

* Windows cmd/powershell 用的都是gbk编码，所以perl的标准输入输出都是gbk编码
* Perl内部的string有所谓的Text String 与Binary String之分，从text到binary的转换为Encode，binary到text的转换为Decode。
* Sptreasheet::XLS读入的每个cell的Value存储的都是Byte String，使用时一定要特别注意。尤其是在匹配的时候，不要用Text String去匹配Byte String，也不要用Byte String去匹配Text String。
* 参见[[http://perldoc.perl.org/perlunitut.html#I%2fO-flow-(the-actual-5-minute-tutorial)|perlunitut]]
* 输入的XLSX表格内容


|   | A      | B      | C      |
| 1 | 大家好 | 我是   | 周杰伦 |
| 2 | 这是   | 第一行 | 第三列 |

== 代码 ==

* 源代码如下，完成的事情就是把一个Excel表格里的每一个Cell的值写入到另外一个Excel表格里，当输入的Excel表格里的某一行第二列包含"我是"这个字符串时，就直接跳过该行。


#!/user/bin/env perl

# Auther: Philip Ye (semimiracle@gmail.com)
# Description: Copy every cell in chs.xlsx to output.xlsx, omit current row if matched some
#   specific chinese character

use strict;
use warnings;
use 5.016;
use Spreadsheet::XLSX;
use Excel::Writer::XLSX;
use Spreadsheet::XLSX;
use Encode;
use utf8;   #To treat Chinese String as Text String, instead of Byte String
binmode STDOUT, ':encoding(gbk)';   #For Chinese output in console

my $excel = Spreadsheet::XLSX->new ('chs.xlsx', undef);
my $wb = Excel::Writer::XLSX->new('output.xlsx');
my $ws = $wb->add_worksheet();

foreach my $sheet (@{$excel->{Worksheet}}) {
  foreach my $row ($sheet->{MinRow} .. $sheet->{MaxRow}) {
    my $flag = decode("utf8", $sheet->{Cells}[$row][1]->{Val});
    say (decode("gbk",encode("gbk",$flag)));
    if (defined ($flag) && $flag =~ /我是/) {     #Target match chinese text string is "我是"
      next;
    }
    foreach my $col ($sheet ->{MinCol} ..  $sheet->{MaxCol}) {
      my $cell = $sheet->{Cells}[$row][$col];
      if ($cell) {
        say (decode("gbk",encode("gbk",decode("utf8",$cell->{Val}))));
        $ws->write($row, $col, decode("utf8",$cell->{Val}));
      }
    }
  }
}




== 解释 ==
* `use utf8`;
* 参见perldoc说明，use utf8之后，源代码内的所有字符都会被perl parse用utf-8编码处理。所以我们在匹配的时候可以直接在匹配目标里使用中文。
* [[http://perldoc.perl.org/utf8.html|utf8 perldoc]]
* `my $excel = Spreadsheet::XLSX->new ('chs.xlsx', undef);`
* 由于chs.xlsx存储的时候就使用的是utf8编码，因此在new的时候，Converter形参并传入的是undef。此处并未使用utf-8 => gbk的Text::Iconv，是因为我们在源代码里处理的时候，除了输出之后，全部使用utf-8，因此没必要转成utf-8。
* `my $flag = decode("utf8", $sheet->{Cells}[$row][1]->{Val});`
* Spreadsheet::XLSX处理时，Cell的Value是以Byte String的方式储存的，并且该Byte String是utf-8的Byte String。由于在代码中`/我是/`是Text String，因此在匹配之前要先把Cell的Value从Byte String转换为Text String.
* `say (decode("gbk",encode("gbk",$flag)));`
* 前面已经说过了，Windows Console输出用的是gbk编码，因此在输出到console的时候，需要完成utf8 => gbk的转换。注意到此时$flag已经是utf-8的Text String了，因此要先把它用encode函数转换为gbk的Byte String，然后用decode函数转换为gbk的Text String，最后将gbk的Text String输出到console。
* `if (defined ($flag) && $flag =~ /我是/) {     #Target match chinese text string is "我是"`
* 注意此处已经是utf8的Text Stringmatch了，因此没必要使用/u的unicode match modifier。(当然加了也不会错:))
* `$ws->write($row, $col, decode("utf8",$cell->{Val}));`
* $cell->{Val}是一个Byte String，在调用Excel::Writer::ELXS的write的时候，需要的是Text String，因此要掉decode把Byte String转换成Text String。
