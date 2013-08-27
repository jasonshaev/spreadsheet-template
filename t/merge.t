#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Data::Dumper;

use Spreadsheet::ParseXLSX;
use Spreadsheet::Template;

my $template = Spreadsheet::Template->new;
my $data = do { local $/; local @ARGV = ('t/data/merge.json'); <> };

{
    my $excel = $template->render(
        $data,
        {
            rows => [
                {
                    value1 => "Merge 1",
                    value2 => "Merge 2",
                    value3 => "Merge 3",
                    value4 => "Merge 4"
                }
            ],
        }
    );

    open my $fh, '<', \$excel;
    my $wb = Spreadsheet::ParseXLSX->new->parse($fh);
    is($wb->worksheet_count, 1);

    my $ws = $wb->worksheet(0);
    is($ws->get_name, 'Merge Report 1');

    # In the template, the 4 columns are merged
    # with contents = "Merged Cells"
    for my $col (0..3) {
        if ($col == 0) {
            is($ws->get_cell(0, $col)->value, 'Merged Cells');
        } else {
            is($ws->get_cell(0, $col)->value, '');
        }
    }
}

done_testing;
