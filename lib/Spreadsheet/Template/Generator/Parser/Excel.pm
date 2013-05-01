package Spreadsheet::Template::Generator::Parser::Excel;
use Moose::Role;

with 'Spreadsheet::Template::Generator::Parser';

requires 'make_excel';

sub parse {
    my $self = shift;
    my ($filename) = @_;

    my $excel = $self->make_excel($filename);
    return $self->_parse_workbook($excel);
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
        name          => $sheet->get_name,
        row_heights   => [ $sheet->get_row_heights ],
        column_widths => [ $sheet->get_col_widths ],
        # XXX Spreadsheet::ParseExcel doesn't currently support extracting the
        # currently selected cells or worksheets
        cells         => [],
    };

    my ($rmin, $rmax) = $sheet->row_range;
    my ($cmin, $cmax) = $sheet->col_range;

    splice @{ $data->{row_heights} }, $rmax + 1;
    splice @{ $data->{column_widths} }, $cmax + 1;

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

    my $contents = $cell->unformatted;
    my $type = $cell->type;
    my $formula = $cell->{Formula}; # XXX

    if ($type eq 'Numeric') {
        $type = 'number';
    }
    elsif ($type eq 'Text') {
        $type = 'string';
    }
    elsif ($type eq 'Date') {
        $type = 'date_time';
    }
    else {
        die "unknown type $type";
    }

    my $data = {
        contents => $self->_filter_cell_contents($contents, $type),
        type     => $type,
        ($formula ? (formula => $formula) : ()),
    };

    return $data;
}

sub _filter_cell_contents {
    my $self = shift;
    my ($contents) = @_;
    return $contents;
}

no Moose::Role;

1;
