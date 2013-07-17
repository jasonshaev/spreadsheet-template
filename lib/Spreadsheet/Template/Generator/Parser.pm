package Spreadsheet::Template::Generator::Parser;
use Moose::Role;
# ABSTRACT: role for classes which parse an existing spreadsheet

requires 'parse';

has filename => (
    is       => 'ro',
    isa      => 'Str',
    required => 1,
);

no Moose::Role;

1;
