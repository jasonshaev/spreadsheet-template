package Spreadsheet::Template::Writer;
use Moose::Role;
# ABSTRACT: role for classes which write spreadsheet files from a template

requires 'write';

no Moose::Role;

1;
