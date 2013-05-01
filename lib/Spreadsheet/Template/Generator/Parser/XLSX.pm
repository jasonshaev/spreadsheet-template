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

    my $book_xml = $self->_parse_xml("xl/workbook.xml");
    $self->_parse_selected_sheet($excel, $book_xml);

    for my $sheet ($excel->worksheets) {
        my $sheet_xml = $self->_parse_xml("xl/$sheet->{path}");

        $self->_parse_cell_sizes($sheet, $sheet_xml);
        $self->_parse_formulas($sheet, $sheet_xml);
        $self->_parse_sheet_selection($sheet, $sheet_xml);
    }
}

sub _parse_selected_sheet {
    my $self = shift;
    my ($excel, $root) = @_;

    my ($node) = $root->find_nodes('//workbookView');
    my $selected = $node->att('activeTab');

    $excel->{SelectedSheet} = defined($selected) ? 0+$selected : 0;
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
        my ($row, $col) = $self->_cell_to_row_col($cell_id);
        my $cell = $sheet->get_cell($row, $col);
        $cell->{Formula} = "=" . $formula->text;
    }
}

sub _parse_sheet_selection {
    my $self = shift;
    my ($sheet, $root) = @_;

    my ($selection) = $root->find_nodes('//selection');
    my $cell = $selection->att('activeCell');

    $sheet->{Selection} = [ $self->_cell_to_row_col($cell) ];
}

sub _parse_xml {
    my $self = shift;
    my ($subfile) = @_;

    my $filename = $self->filename;

    my $zip = Archive::Zip->new;
    die "Can't open $filename as zip file"
        unless $zip->read($filename) == Archive::Zip::AZ_OK;

    my $contents = $zip->memberNamed($subfile)->contents;
    next unless $contents;

    my $xml = XML::Twig->new;
    $xml->parse($contents);

    return $xml->root;
}

sub _cell_to_row_col {
    my $self = shift;
    my ($cell) = @_;

    my ($col, $row) = $cell =~ /([A-Z]+)([0-9]+)/;
    $col =~ tr/A-Z/0-9A-P/;
    $col = POSIX::strtol($col, 26);
    $row = $row - 1;

    return ($row, $col);
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
