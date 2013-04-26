package Spreadsheet::Template::Writer::Excel;
use Moose::Role;

use Class::Load 'load_class';

with 'Spreadsheet::Template::Writer';

requires 'excel_class';

has excel => (
    is      => 'ro',
    isa     => 'Object',
    lazy    => 1,
    default => sub {
        my $self = shift;
        load_class($self->excel_class);
        $self->excel_class->new($self->_fh);
    },
);

has _fh => (
    is      => 'ro',
    isa     => 'FileHandle',
    lazy    => 1,
    default => sub {
        my $self = shift;
        open my $fh, '>', $self->_output
            or die "Failed to open filehandle: $!";
        binmode $fh;
        return $fh;
    },
);

has _output => (
    is      => 'ro',
    isa     => 'ScalarRef[Maybe[Str]]',
    default => sub { \(my $str) },
);

sub write {
    my $self = shift;

    # ...

    $self->excel->close;
    return ${ $self->_output };
}

no Moose::Role;

1;
