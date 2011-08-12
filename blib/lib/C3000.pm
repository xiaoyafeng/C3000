package C3000;
use warnings;
use strict;
use Readonly;
use Win32::OLE;
use Win32::OLE::Variant;
use Carp;
use Encode;
use DateTime;
use DateTime::Format::Natural;
use DBI;
Readonly::Scalar my $DEBUG => 0;
use constant cnNoFlags => 0;
use constant cnNoCheckpoint => 0;


=head1 NAME

C3000 - A perl wrap for C3000 API
	This is a simple wrap of C3000 API. For more details, please refer to: 
	meter2cash CONVERGE Business Objects Component Interfaces User Guide 

=head1 VERSION

Version 0.01

=cut

our $VERSION = '0.01';

=head1 SYNOPSIS
 
   use C3000;
   my $hl = C3000->new();
   my $value = $hl->accu_LP('ADAS_VAL_RAW', '%', '+A', 'yesterday', 'today');
   my $value = $hl->get_single_LP('ADAS_VAL_RAW', '%', '+A', 'yesterday');



=head1 Methods 

=head2 new

init sub

=cut

my $Empty;

sub new {
    my $this              = shift;
	
    #init start
    #my $dbh = DBI->connect("dbi:ADO:Provider=OraOLEDB.Oracle;Data Source=athenaDB;", 'athena', 'athena', ) or croak $DBI::errstr;

    my $container_manager = Win32::OLE->new('Athena.CT.ContainerManager.1');
    my $active_element_manager =
      Win32::OLE->new('Athena.CT.ActiveElementManager.1');
    my $active_element_template_manager =
      Win32::OLE->new('Athena.CT.ActiveElementTemplateManager.1');
    my $security_token = Win32::OLE->new('AthenaSecurity.UserSessions.1');
    $security_token->ConvergeLogin( 'Administrator', 'Athena', 0, 666 );
    my $adas_instance_manager =
      Win32::OLE->new('Athena.AS.ADASInstanceManager.1')
      or croak;
    my $adas_template_manager =
      Win32::OLE->new('Athena.AS.AdasTemplateManager.1')
      or croak;
    my $container_template_manager =
      Win32::OLE->new('Athena.CT.ContainerTemplateManager.1')
      or croak;
    my $data_segment_helper = Win32::OLE->new('AthenaSecurity.DataSegment.1')
      or croak;
    my $meter_data_interface =
      Win32::OLE->new('DeviceAndMeterdata.ADASDeviceAndMeterdata.1');
      #init end
    

    my $self = {
	    #dbh			     => $dbh,
        MeterDataInterface           => $meter_data_interface,
        ContainerManager             => $container_manager,
        ContainerTemplateManager     => $container_template_manager,
        ActiveElementManager         => $active_element_manager,
        ActiveElementTemplateManager => $active_element_template_manager,
        UserSessions                 => $security_token,
        ADASInstanceManager          => $adas_instance_manager,
        ADASTemplateManager          => $adas_template_manager,
        DataSegment                  => $data_segment_helper,
    };
    return bless $self, $this;

}

=head2 get_LP


	my $rs = get_LP(....);
	while(!$rs->EOF){ 
		...
	$rs->MoveNext();
	}


return a recordset of LoadProfile
low function, just ignore it.



=cut

sub get_LP {
    my $self = shift;
    my ( $return_fields, $criteria ) = @_;
    my $rs =
      $self->{'MeterDataInterface'}
      ->GetLoadProfile( $return_fields, $criteria, 0 )
      or croak $self->{'MeterDataInterface'}->LastError();
    return $rs;
}

=head2 accu_LP

		return a accumulation value of LoadProfile
	
		$hl->accu_LP('ADAS_VAL_RAW', '海电%', '+A', 'yesterday', 'today');
	
		$hl->accu_LP('ADAS_VAL_NORM', '%', '+A', '2011/1/1 20:00', '2011/1/1 21:00);



=cut

sub accu_LP {
    my $self = shift;
    my ( $LP_return_fields, $device, $var, $date_from, $date_to ) = @_;

     
    

    #get device
    my $device_return_fields = [
        'ADAS_DEVICE',      'ADAS_VARIABLE',
        'ADAS_DEVICE_NAME', 'ADAS_VARIABLE_NAME'
    ];
    my $device_criteria =
      [ [ 'ADAS_DEVICE_NAME', $device ], [ 'ADAS_VARIABLE_NAME', $var ] ];
    my $rs_device =
      $self->search_device( $device_return_fields, $device_criteria );



    $LP_return_fields = [ $LP_return_fields, 'ADAS_TIME_GMT' ];
    $date_from        = $self->convert_VT_DATE($date_from);
    $date_to          = $self->convert_VT_DATE($date_to);
    my $value = 0;

    while ( !$rs_device->EOF ) {
        my $device_id   = $rs_device->Fields('ADAS_DEVICE')->{Value};
        my $var_id      = $rs_device->Fields('ADAS_VARIABLE')->{Value};
        my $LP_criteria = [
            [ 'ADAS_DEVICE',   $device_id ],
            [ 'ADAS_VARIABLE', $var_id ],
            [ 'ADAS_TIME_GMT', $date_from, '>' ],
            [ 'ADAS_TIME_GMT', $date_to,   '<=' ]
        ];
        my $rs =
          $self->{'MeterDataInterface'}
          ->GetLoadProfile( $LP_return_fields, $LP_criteria, 0 )
          or croak $self->{'MeterDataInterface'}->LastError();
        while ( !$rs->EOF ) {
            $value += $rs->Fields( $LP_return_fields->[0] )->{Value};

            carp $rs->Fields( $LP_return_fields->[1] )->{Value}
              . "meter value is"
              . $value
              if $DEBUG == 1;
            $rs->MoveNext();
        }
        $rs_device->MoveNext();
    }
    return $value;
}

=head2 get_single_LP

		return a single value or a hash 


		get_single_LP('ADAS_VAL_RAW', '海电_主表110', '+A', 'today');

		get_single_LP('ADAS_VAL_NORM', '主表%', 'yesterday');          

as above, get_single_LP function return a value  or a hash when value is more than one.

=cut

sub get_single_LP {
    my $self = shift;
    my ( $LP_return_fields, $device, $var, $date ) = @_;
    $LP_return_fields = [$LP_return_fields];
    my $parser = DateTime::Format::Natural->new(
        lang      => 'en',
        format    => 'yyyy/mm/dd',
        time_zone => 'Asia/Taipei',
    );
    my $dt = $parser->parse_datetime($date);
    $dt->subtract( minutes => 1 );
    $dt->set_time_zone('UTC');
    my $date_from            = Variant( VT_DATE, 25569 + $dt->epoch / 86400 );
    my $date_to              = $self->convert_VT_DATE($date);
    my $device_return_fields = [
        'ADAS_DEVICE',      'ADAS_VARIABLE',
        'ADAS_DEVICE_NAME', 'ADAS_VARIABLE_NAME'
    ];
    my $device_criteria =
      [ [ 'ADAS_DEVICE_NAME', $device ], [ 'ADAS_VARIABLE_NAME', $var ] ];
    my $rs_device =
      $self->search_device( $device_return_fields, $device_criteria )
      ;    #record set of device;
    my %LP_values ;
    my $LP_key;
    while ( !$rs_device->EOF ) {
        my $device_id   = $rs_device->Fields('ADAS_DEVICE')->{Value};
        my $var_id      = $rs_device->Fields('ADAS_VARIABLE')->{Value};
        my $LP_criteria = [
            [ 'ADAS_DEVICE',   $device_id ],
            [ 'ADAS_VARIABLE', $var_id ],
            [ 'ADAS_TIME_GMT', $date_from, '>' ],
            [ 'ADAS_TIME_GMT', $date_to,   '<=' ]
        ];
        my $rs =
          $self->{'MeterDataInterface'}
          ->GetLoadProfile( $LP_return_fields, $LP_criteria, 0 )
          or croak $self->{'MeterDataInterface'}->LastError();
        while ( !$rs->EOF ) {
             $LP_key =
                $rs_device->Fields('ADAS_DEVICE_NAME') . "_"
              . $rs_device->Fields('ADAS_VARIABLE_NAME');
            $LP_values{$LP_key} =
              $rs->Fields( $LP_return_fields->[0] )->{Value};
            $rs->MoveNext();
        }
        $rs_device->MoveNext();
    }

    return keys %LP_values <= 1 ? $LP_values{ $LP_key }   :  \%LP_values;

}

=head2 search_device

		return a recordset of device info fitting criteria 
		low function, please refer test file for more details.



=cut

sub search_device {
    my $self = shift;
    my ( $return_fields, $criteria ) = @_;
    my $rs =
      $self->{'MeterDataInterface'}
      ->FindVariable( $return_fields, $criteria, $Empty, 0 )
      or croak $self->{'MeterDataInterface'}->LastError();

    return $rs;
}

=head2  convert_VT_DATE 


C3000 utils, pass to a DateTime obj and return a VT_DATE variable.

=cut

sub convert_VT_DATE {
    shift;
    Readonly::Scalar my $EPOCH       => 25569;
    Readonly::Scalar my $SEC_PER_DAY => 86400;
    my $parser = DateTime::Format::Natural->new(
        lang      => 'en',
        format    => 'yyyy/mm/dd',
        time_zone => 'Asia/Taipei',
    );
    my $date_string = shift;
     return  Variant( VT_DATE, $EPOCH + $date_string->epoch / $SEC_PER_DAY ) if ref($date_string) eq 'DateTime';
    my $dt          = $parser->parse_datetime($date_string);
    $dt->set_time_zone('UTC');
    return Variant( VT_DATE, $EPOCH + $dt->epoch / $SEC_PER_DAY );
}



sub CreateMasterAccount {
	my ($self, $pnParentContainerID, $pnTemplateID, $pnDataSegmentID, $paValues) = @_;

	my $nContainerID;
	my $rsInitValues;
	my @arrayRf;
	for my $nIndex (0 .. $#$paValues){
		$arrayRf[ $nIndex ] = $paValues->[ $nIndex ][0];
	}
	$rsInitValues = $self->{ 'ContainerTemplateManager' }->GetValues($pnTemplateID, \@arrayRf, cnNoFlags, cnNoCheckpoint, $self->{ 'UserSessions' });
	for my $nIndex (0 .. $#$paValues){
		$rsInitValues->Fields($paValues->[ $nIndex ][0])->{'Value'} = $paValues->[ $nIndex ][1];
		}
	$nContainerID = $self->{ 'ContainerManager' }->CreateContainer($pnTemplateID, $self->{ 'ContainerManager' }->GetTopNodeID( cnNoCheckpoint, $self->{ 'UserSessions' }), $rsInitValues, $self->{ 'UserSessions' });
        $self->{ 'ContainerManager' }->MoveContainer($nContainerID, $pnParentContainerID, $self->{ 'UserSessions' });
	$self->{ 'ContainerManager' }->setDataSegment($nContainerID, $pnDataSegmentID, $self->{ 'UserSessions' });
	return $nContainerID;
}

sub GetContainerID {
	my $self       = shift;
	my $paCriteria = shift;
	my $rs;
	$rs = $self->{ 'ContainerManager' }->Search(['ContainerID'], $paCriteria, $Empty, cnNoFlags, cnNoCheckpoint, $self->{ 'UserSessions' });
	return $rs->Fields('ContainerID')->{'Value'};
}

sub GetDataSegmentID {
	my $self	      = shift;
	my $psDataSegmentName = shift;
        my $rs;
	$rs = $self->{ 'DataSegment' }->GetDataSegmentInfo($self->{ 'UserSessions' }->GetInstallationID(), ['DataSegNb'], [['DataSegName', $psDataSegmentName]]);
	return $rs->Fields('DataSegNb')->{'Value'};

}

sub GetTemplateID {
	my $self           = shift;
	my $psTemplatename = shift;
	my ($arrayRf, $arrayCr);
	my $rsTemplate;
	$arrayRf = ['ContainerTemplateID'];
	$arrayCr = [['ContainerTemplateName', $psTemplatename]];
	$rsTemplate = $self->{ 'ContainerTemplateManager' }->Search($arrayRf, $arrayCr, $Empty, cnNoFlags, cnNoCheckpoint, $self->{ 'UserSessions' });
	return $rsTemplate->Fields('ContainerTemplateID')->{'Value'};


}


=head1 AUTHOR

Andy Xiao, C<< <xyf.gmail.com> >>

=head1 BUGS

Please report any bugs or feature requests to C<bug-c3000 at rt.cpan.org>, or through
the web interface at L<http://rt.cpan.org/NoAuth/ReportBug.html?Queue=C3000>.  I will be notified, and then you'll
automatically be notified of progress on your bug as I make changes.




=head1 TODO 
extends qw(
	C3000::MeterDataIntergface
	C3000::ADASInstanceManager
	C3000::ContainerManager
	C3000::ContainerTemplateManager
	C3000::ActiveElementManager
	C3000::ActiveElementTemplateManager
	C3000::SecurityUserSession
	C3000::DataSegment

);
consider use Moose

=over 4

=item * RT: CPAN's request tracker (report bugs here)

L<http://rt.cpan.org/NoAuth/Bugs.html?Dist=C3000>

=item * AnnoCPAN: Annotated CPAN documentation

L<http://annocpan.org/dist/C3000>

=item * CPAN Ratings

L<http://cpanratings.perl.org/d/C3000>

=item * Search CPAN

L<http://search.cpan.org/dist/C3000/>

=back


=head1 ACKNOWLEDGEMENTS


=head1 LICENSE AND COPYRIGHT

Copyright 2011 Andy Lester.

This program is free software; you can redistribute it and/or modify it
under the terms of either: the GNU General Public License as published
by the Free Software Foundation; or the Artistic License.

See http://dev.perl.org/licenses/ for more information.


=cut

1;    # End of C3000
