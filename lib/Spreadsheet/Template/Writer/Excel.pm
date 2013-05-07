package Spreadsheet::Template::Writer::Excel;
use Moose::Role;

use Class::Load 'load_class';
use List::Util 'first';

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

has _colors => (
    is      => 'ro',
    isa     => 'HashRef[Int]',
    default => sub {
        {
            black   => 8,
            blue    => 12,
            brown   => 16,
            cyan    => 15,
            gray    => 23,
            green   => 17,
            lime    => 11,
            magenta => 14,
            navy    => 18,
            orange  => 53,
            pink    => 33,
            purple  => 20,
            red     => 10,
            silver  => 22,
            white   => 9,
            yellow  => 13,

        }
    },
);

has _formats => (
    is      => 'ro',
    isa     => 'HashRef',
    default => sub { {} },
);

sub write {
    my $self = shift;
    my ($data) = @_;

    $self->_write_workbook($data);

    $self->excel->close;
    return ${ $self->_output };
}

sub _write_workbook {
    my $self = shift;
    my ($data) = @_;

    # XXX no way to write default cell properties

    if (exists $data->{properties}) {
        $self->excel->set_properties(%{ $data->{properties} });
    }

    for my $sheet (@{ $data->{worksheets} }) {
        $self->_write_worksheet($sheet);
    }

    if (exists $data->{selected}) {
        $self->excel->sheets($data->{selected})->activate;
    }
}

sub _write_worksheet {
    my $self = shift;
    my ($data) = @_;

    my $sheet = $self->excel->add_worksheet(
        exists $data->{name} ? ($data->{name}) : ()
    );

    if (exists $data->{tab_color}) {
        $sheet->set_tab_color($self->_color($data->{tab_color}));
    }

    if (exists $data->{zoom}) {
        $sheet->set_zoom($data->{zoom});
    }

    if (exists $data->{hidden}) {
        # XXX this won't work on the first worksheet, since you can't hide
        # active worksheets - need to restructure things a bit to fix this
        $sheet->hide;
    }

    if (exists $data->{selected}) {
        $sheet->set_selection(@{ $data->{selected} });
    }

    if (exists $data->{freeze}) {
        $sheet->freeze_panes(@{ $data->{freeze} });
    }

    if (exists $data->{split}) {
        $sheet->split_panes(@{ $data->{split} });
    }

    if (exists $data->{column_widths}) {
        for my $i (0..$#{ $data->{column_widths} }) {
            # XXX hidden columns?
            $sheet->set_column($i, $i, $data->{column_widths}[$i]);
        }
    }

    if (exists $data->{row_heights}) {
        for my $i (0..$#{ $data->{row_heights} }) {
            # XXX hidden rows?
            $sheet->set_row($i, $data->{row_heights}[$i]);
        }
    }

    for my $row (0..$#{ $data->{cells} }) {
        for my $col (0..$#{ $data->{cells}[$row] }) {
            $self->_write_cell($data->{cells}[$row][$col], $sheet, $row, $col);
        }
    }
}

sub _write_cell {
    my $self = shift;
    my ($data, $sheet, $row, $col) = @_;

    my $write_method = 'write';
    if (exists $data->{type}) {
        $write_method = "write_$data->{type}";
    }

    my $format;
    if (exists $data->{format}) {
        my $properties = {
            map {
                my $v = $data->{format}{$_};
                $_ => JSON::is_bool($v) ? ($v ? 1 : 0)
                    : $_ =~ /color/     ? $self->_color($v)
                    :                     $v
            } keys %{ $data->{format} }
        };
        $format = $self->_format($properties);
    }

    # XXX handle merged cells

    $sheet->$write_method(
        $row, $col,
        $data->{contents},
        (defined $format ? ($format) : ()),
    );
}

sub _color {
    my $self = shift;
    my ($color) = @_;

    if (exists $self->_colors->{$color}) {
        return $self->_colors->{$color};
    }
    else {
        my $hex = qr/[0-9a-fA-F]/;
        my ($r, $g, $b) = $color =~ /^#($hex$hex)($hex$hex)($hex$hex)$/;

        my %used_colors = reverse %{ $self->_colors };
        my $new_idx = first { !exists $used_colors{$_} } 8..63;
        die "too many colors" unless defined $new_idx;

        $self->excel->set_custom_color(
            $new_idx,
            map { oct("0x$_") } $r, $g, $b
        );
        $self->_colors->{$color} = $new_idx;

        return $new_idx;
    }
}

sub _format {
    my $self = shift;
    my ($format_properties) = @_;

    my $key = JSON->new->canonical->encode($format_properties);
    if (exists $self->_formats->{$key}) {
        return $self->_formats->{$key};
    }
    else {
        my $format = $self->excel->add_format(%$format_properties);
        $self->_formats->{$key} = $format;
        return $format;
    }
}

no Moose::Role;

1;
