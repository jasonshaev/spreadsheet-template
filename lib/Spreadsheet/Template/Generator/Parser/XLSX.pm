package Spreadsheet::Template::Generator::Parser::XLSX;
use Moose;

use XML::Entities;
use Spreadsheet::XLSX;

with 'Spreadsheet::Template::Generator::Parser';

sub parse {
    my $self = shift;
    my ($filename) = @_;

    my $excel = Spreadsheet::XLSX->new($filename);
    return $self->_parse_workbook($excel);
    my $data = {
        worksheets => [],
    };
    for my $sheet ($excel->worksheets) {
        push @{ $data->{worksheets} }, $self->_parse
    }
}

sub _parse_workbook {
    my $self = shift;
    my ($excel) = @_;

    my $data = {
        worksheets => [],
    };

    for my $sheet ($excel->worksheets) {
        push @{ $data->{worksheets} }, $self->_parse_worksheet($sheet);
    }

    return $data;
}

sub _parse_worksheet {
    my $self = shift;
    my ($sheet) = @_;

    my $data = {
        cells => [],
    };

    my ($rmin, $rmax) = $sheet->row_range;
    my ($cmin, $cmax) = $sheet->col_range;

    for my $row (0..$rmin - 1) {
        push @{ $data->{cells} }, [];
    }

    for my $row ($rmin..$rmax) {
        my $row_data = [];
        for my $col (0..$cmin - 1) {
            push @$row_data, {};
        }
        for my $col ($cmin..$cmax) {
            if (my $cell = $sheet->get_cell($row, $col)) {
                push @$row_data, $self->_parse_cell($cell);
            }
            else {
                push @$row_data, {};
            }
        }
        push @{ $data->{cells} }, $row_data
    }

    return $data;
}

sub _parse_cell {
    my $self = shift;
    my ($cell) = @_;


    my $data = {
        # XXX this decode call really feels like a bug in Spreadsheet::XLSX
        contents => XML::Entities::decode('all', $cell->unformatted),
        type     => $cell->type,
    };

    return $data;
}

__PACKAGE__->meta->make_immutable;
no Moose;

1;
