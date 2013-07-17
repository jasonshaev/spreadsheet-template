package Spreadsheet::Template::Generator::Parser;
use Moose::Role;
# ABSTRACT: role for classes which parse an existing spreadsheet

requires 'parse';

no Moose::Role;

1;
