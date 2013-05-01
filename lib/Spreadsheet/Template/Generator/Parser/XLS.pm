package Spreadsheet::Template::Generator::Parser::XLS;
use Moose;

use Spreadsheet::ParseExcel;

with 'Spreadsheet::Template::Generator::Parser::Excel';

sub make_excel {
    my $self = shift;
    my ($filename) = @_;

    my $parser = Spreadsheet::ParseExcel->new;
    my $excel = $parser->parse($filename);
    die $parser->error unless $excel;

    # just for consistency
    for my $sheet ($excel->worksheets) {
        $sheet->{RowHeight} = [
            map { defined $_ ? $_ : $sheet->get_default_row_height }
                $sheet->get_row_heights
        ];
        $sheet->{ColWidth} = [
            map { defined $_ ? $_ : $sheet->get_default_col_width }
                $sheet->get_col_widths
        ];
    }

    # XXX no formula support yet

    return $excel;
}

__PACKAGE__->meta->make_immutable;
no Moose;

1;
