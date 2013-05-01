package Spreadsheet::Template::Generator::Parser;
use Moose::Role;

requires 'parse';

has filename => (
    is       => 'ro',
    isa      => 'Str',
    required => 1,
);

no Moose::Role;

1;
