package Spreadsheet::Template::Generator::Parser::XLSX;
use Moose;

use Spreadsheet::XLSX;
use XML::Entities;
use XML::Twig;

with 'Spreadsheet::Template::Generator::Parser::Excel';

sub make_excel {
    my $self = shift;
    my ($filename) = @_;
    my $excel = Spreadsheet::XLSX->new($filename);

    # XXX Spreadsheet::XLSX doesn't extract this information currently
    my $zip = Archive::Zip->new;
    die "Can't open $filename as zip file"
        unless $zip->read($filename) == Archive::Zip::AZ_OK;

    for my $sheet ($excel->worksheets) {
        my $contents = $zip->memberNamed("xl/$sheet->{path}")->contents;
        next unless $contents;

        my @column_widths;
        my @row_heights;

        my $xml = XML::Twig->new;
        $xml->parse($contents);
        my $root = $xml->root;

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

    return $excel;
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
