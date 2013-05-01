package Spreadsheet::Template::Generator::Parser::XLS;
use Moose;

use Spreadsheet::ParseExcel;

with 'Spreadsheet::Template::Generator::Parser::Excel';

sub _build_excel {
    my $self = shift;

    my $parser = Spreadsheet::ParseExcel->new;
    my $excel = $parser->parse($self->filename);
    die $parser->error unless $excel;

    $self->_fixup_excel($excel);

    return $excel;
}

sub _fixup_excel {
    my $self = shift;
    my ($excel) = @_;

    for my $sheet ($excel->worksheets) {
        $self->_normalize_cell_sizes($sheet);
        $self->_parse_formulas($sheet);
    }
}

sub _normalize_cell_sizes {
    my $self = shift;
    my ($sheet) = @_;

    # just for consistency
    $sheet->{RowHeight} = [
        map { defined $_ ? $_ : $sheet->get_default_row_height }
            $sheet->get_row_heights
    ];
    $sheet->{ColWidth} = [
        map { defined $_ ? $_ : $sheet->get_default_col_width }
            $sheet->get_col_widths
    ];
}

sub _parse_formulas {
    # XXX no formula support yet
}

__PACKAGE__->meta->make_immutable;
no Moose;

1;
