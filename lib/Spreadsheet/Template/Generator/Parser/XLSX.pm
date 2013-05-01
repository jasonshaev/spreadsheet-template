package Spreadsheet::Template::Generator::Parser::XLSX;
use Moose;

use POSIX;
use Spreadsheet::XLSX;
use XML::Entities;
use XML::Twig;

with 'Spreadsheet::Template::Generator::Parser::Excel';

sub _build_excel {
    my $self = shift;
    my $excel = Spreadsheet::XLSX->new($self->filename);
    $self->_fixup_excel($excel);
    return $excel;
}

# XXX Spreadsheet::XLSX doesn't extract this information currently
sub _fixup_excel {
    my $self = shift;
    my ($excel) = @_;

    my $filename = $self->filename;

    my $zip = Archive::Zip->new;
    die "Can't open $filename as zip file"
        unless $zip->read($filename) == Archive::Zip::AZ_OK;

    for my $sheet ($excel->worksheets) {
        my $contents = $zip->memberNamed("xl/$sheet->{path}")->contents;
        next unless $contents;

        my $xml = XML::Twig->new;
        $xml->parse($contents);
        my $root = $xml->root;

        $self->_parse_cell_sizes($sheet, $root);
        $self->_parse_formulas($sheet, $root);
    }
}

sub _parse_cell_sizes {
    my $self = shift;
    my ($sheet, $root) = @_;

    my @column_widths;
    my @row_heights;

    my ($format) = $root->find_nodes('//sheetFormatPr');
    my $default_row_height = $format->att('defaultRowHeight');
    my $default_column_width = $format->att('baseColWidth');

    for my $col ($root->find_nodes('//col')) {
        $column_widths[$col->att('min') - 1] = $col->att('width');
    }

    for my $row ($root->find_nodes('//row')) {
        $row_heights[$row->att('r') - 1] = $row->att('ht');
    }

    $sheet->{DefRowHeight} = 0+$default_row_height;
    $sheet->{DefColWidth} = 0+$default_column_width;
    $sheet->{RowHeight} = [
        map { defined $_ ? 0+$_ : 0+$default_row_height } @row_heights
    ];
    $sheet->{ColWidth} = [
        map { defined $_ ? 0+$_ : 0+$default_column_width } @column_widths
    ];
}

sub _parse_formulas {
    my $self = shift;
    my ($sheet, $root) = @_;

    for my $formula ($root->find_nodes('//f')) {
        my $cell_id = $formula->parent->att('r');
        my ($col, $row) = $cell_id =~ /([A-Z]+)([0-9]+)/;
        $col =~ tr/A-Z/0-9A-P/;
        $col = POSIX::strtol($col, 26);
        $row = $row - 1;
        my $cell = $sheet->get_cell($row, $col);
        $cell->{Formula} = "=" . $formula->text;
    }
}

# XXX this stuff all feels like working around bugs in Spreadsheet::XLSX -
# maybe look into that at some point
sub _filter_cell_contents {
    my $self = shift;
    my ($contents, $type) = @_;

    $contents = XML::Entities::decode('all', $contents);

    if ($type eq 'number') {
        $contents = 0+$contents;
    }

    return $contents;
}

__PACKAGE__->meta->make_immutable;
no Moose;

1;
