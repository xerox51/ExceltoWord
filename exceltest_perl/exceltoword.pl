use Spreadsheet::XLSX;
use Win32::Word::Writer;
use Encode;

sub word{
    my($item,$filename) = @_;
    my $oWriter = Win32::Word::Writer->new();
    my @array = split(/\n/,$item);
    foreach $k (@array){
        $oWriter->Write(join "",$k,"\n");
    }
    $oWriter->SaveAs($filename);
}

sub removewhitespace {
    my($item) = @_;
    $item =~ s/^\s+|\s+$//g;
    return $item;
}

sub joinlist{
    my(@item) = @_;
    my $length = @item;
    my @letters = ('B','C','D','E','F','G','H');
    $item[0] = join "",removewhitespace($item[0]),"\n";
    if((defined $item[1]) && (!removewhitespace($item[1]))) {
        $item[1] = join "","A.",removewhitespace($item[1]);
    }
    else {
        $item[1] = "";
    }
    if ($length == 3){
        $item[$length - 1] = join ""," 答案：",removewhitespace($item[$length - 1]),"\n";
        return join("",@item);
    }
    if ($length > 3){
        for($i = 0;$i < $length - 3;$i++){
            if((defined $item[$i+2]) && (!removewhitespace($item[$i+2]))) {
                $item[$i+2] = join ""," ",$letters[$i],".",removewhitespace($item[$i+2]);
            }
            else{
                $item[$i+2] = "";
            }
        }
        $item[$length - 1] = join ""," 答案：",removewhitespace($item[$length - 1]),"\n";
        return join("",@item);
    }
}

sub process{
    my(@item) = @_;
    my @questionindex = ();
    my @question = ();
    my @answer = ();
    my @choice = ();
    my $size = @item;
    for(my $i = 0;$i < $size;$i++){
        if ($item[$i] =~ /题目/){
            push(@question,$i);
        }
        if ($item[$i] =~ /[A-H]/i){
            push(@choice,$i);
        }
        if ($item[$i] =~ /答案/){
            push(@answer,$i);
        }
    }
    if(@question){
        push(@questionindex,$question[0]);
    }
    if(@choice){
        push @questionindex,@choice;
    }
    if(@answer){
        push (@questionindex,$answer[0]);
    }
    return @questionindex;
}

my $excel = Spreadsheet::XLSX->new('safety.xlsx');
my @temp = ();
my $size = undef;

foreach my $sheet (@{$excel->{Worksheet}}) {
    my @sum = ();
    $sheet->{MaxRow} ||= $sheet->{MinRow};
    foreach my $row ($sheet->{MinRow} .. $sheet->{MaxRow}) {  
        $sheet->{MaxCol} ||= $sheet->{MinCol};
        my $everyrowlist = "";
        if($row == 0){
            my @lineone = ();
            foreach my $col ($sheet->{MinCol} ..  $sheet->{MaxCol}) {
                my $cell = $sheet->{Cells}[$row][$col];
                push(@lineone,$cell->{Val});
            }
            @temp = process(@lineone);
            $size = @temp;
        }
        if($row > 0){
            my @tempeveryrow = ();
            foreach my $col ($sheet->{MinCol} ..  $sheet->{MaxCol}) {
                if (grep {$col eq $_ } @temp){
                    my $cell = $sheet->{Cells}[$row][$col];
                    push(@tempeveryrow,$cell->{Val});
                }
            }
            if(@tempeveryrow){
                $everyrowlist = join "","(",$row,")",joinlist(@tempeveryrow);
                push(@sum,$everyrowlist);
            }
        }
    }
    if(@sum){
    word(join("",@sum),join("",$sheet->{Name},"Perl.doc"));
    }
}