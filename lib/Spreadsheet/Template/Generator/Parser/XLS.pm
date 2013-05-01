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

    $self->_parse_selected_sheet($excel);

    for my $sheet ($excel->worksheets) {
        $self->_normalize_cell_sizes($sheet);
        $self->_parse_formulas($sheet);
        $self->_parse_selection($sheet);
    }
}

sub _parse_selected_sheet {
    my $self = shift;
    my ($excel) = @_;
    # XXX no selected sheet support yet
    $excel->{SelectedSheet} = 0;
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

sub _parse_selection {
    my $self = shift;
    my ($sheet) = @_;
    # XXX no selection support yet
    $sheet->{Selection} = [ 0, 0 ];
}

__PACKAGE__->meta->make_immutable;
no Moose;

1;
