package Spreadsheet::Template::Processor;
use Moose::Role;
# ABSTRACT: role for classes which preprocess a template file before rendering

requires 'process';

no Moose::Role;

1;
