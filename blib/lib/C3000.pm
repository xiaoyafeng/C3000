package C3000;

use Win32::OLE;
use Win32::OLE::Variant;
use Carp;
use Encode;
use DateTime;
use DateTime::Format::Natural;
use constant DEBUG => 0;
#
=head1
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
=head1 NAME

C3000 - A perl wrap for C3000 API

=head1 VERSION

Version 0.01

=cut

our $VERSION = '0.01';
 
=head1 SYNOPSIS
meter2cash CONVERGE 
 
Business Objects Component Interfaces 
 
User Guide
This is a simple wrap of C3000, for more details, please refer to: 
meter2cash CONVERGE Business Objects Component Interfaces User Guide 

code snippet:
   use C3000;
    my $meter = '%';
my $meter_var  = '+A';
my $return_fields = ['ADAS_DEVICE', 'ADAS_VARIABLE', 'ADAS_DEVICE_NAME','ADAS_VARIABLE_NAME'];
my $criteria = [['ADAS_DEVICE_NAME',  $meter], ['ADAS_VARIABLE_NAME', $meter_var]];
my $rs_device;
my $rs_LP;
my $rs_device = $hl->search_device($return_fields, $criteria);  #record set of device;

$return_fields = [ 'ADAS_TIME_GMT', 'ADAS_VAL_RAW', 'ADAS_USER_STATUS'];
my $device_id = $rs_device->Fields('ADAS_DEVICE')->{Value};
my $variable_id = $rs_device->Fields('ADAS_VARIABLE')->{Value};
my $date_from  = $hl->convert_VT_DATE('yesterday');
my $date_to    = $hl->convert_VT_DATE('today');
$criteria = [['ADAS_DEVICE', $device_id], ['ADAS_VARIABLE', $variable_id], ['ADAS_TIME_GMT', $date_from, '>'], ['ADAS_TIME_GMT', $date_to, '<=']];

$rs_LP = $hl->get_LP($return_fields, $criteria);


=head1 EXPORT

A list of functions that can be exported.  You can delete this section
if you don't export anything, such as for a purely object-oriented module.


=head2  new
init sub 

=cut
my $Empty;
sub new {
	my $this = shift;
	my $container_manager = Win32::OLE->new('Athena.CT.ContainerManager.1');
	my $active_element_manager = Win32::OLE->new('Athena.CT.ActiveElementManager.1');
	my $active_element_template_manager = Win32::OLE->new('Athena.CT.ActiveElementTemplateManager.1');
	my $security_token = Win32::OLE->new('AthenaSecurity.UserSessions.1');
        	$security_token->ConvergeLogin('Administrator', 'Athena', 0, 666);
	my $adas_instance_manager = Win32::OLE->new('Athena.AS.ADASInstanceManager.1') or die;
	my $adas_template_manager = Win32::OLE->new('Athena.AS.AdasTemplateManager.1') or die;
	my $container_template_manager = Win32::OLE->new('Athena.CT.ContainerTemplateManager.1') or die;
	my $data_segment_helper = Win32::OLE->new('AthenaSecurity.DataSegment.1') or die;
	my $meter_data_interface = Win32::OLE->new('DeviceAndMeterdata.ADASDeviceAndMeterdata.1');



	my $self = { 
			MeterDataInterface 	   	=> 	 $meter_data_interface,
			ContainerManager           	=>	 $container_manager,
			ContainerTemplateManager	=>	 $container_template_manager,
			ActiveElementManager       	=>	 $active_element_manager,
			ActiveElementTemplateManager  	=>	 $active_element_template_manager,
		        UserSessions			=>	 $security_token,
			ADASInstanceManager		=> 	 $adas_instance_manager,
			ADASTemplateManager		=>	 $adas_template_manager,
			DataSegment			=> 	 $data_segment_helper,			
			};
	return bless $self, $this;

}

=head2 get_LP
return a recordset of LoadProfile


e.g.:
while(!$rs->EOF){
...
$rs->MoveNext();
}
=cut
sub get_LP{
	my $self = shift;
	my ($return_fields, $criteria) = @_;
	my $rs = $self->{'MeterDataInterface'}->GetLoadProfile($return_fields, $criteria, 0) or die $hl->{'MeterDataInterface'}->LastError();
	return $rs;	
}

=head2 accu_LP
return a accumulation value of LoadProfile
$hl->accu_LP('ADAS_VAL_RAW', '海电%', '+A', 'yesterday', 'today');
$hl->accu_LP('ADAS_VAL_NORM', '%', '+A', '2011/1/1 20:00', '2011/1/1 21:00);
=cut
sub accu_LP{
	my $self = shift;
	my ($LP_return_fields, $device, $var, $date_from, $date_to) = @_;
	$LP_return_fields = [ $LP_return_fields, 'ADAS_TIME_GMT' ];
	$date_from = $self->convert_VT_DATE($date_from);
	$date_to   = $self->convert_VT_DATE($date_to);
	my $device_return_fields = ['ADAS_DEVICE', 'ADAS_VARIABLE', 'ADAS_DEVICE_NAME','ADAS_VARIABLE_NAME'];
	my $device_criteria = [['ADAS_DEVICE_NAME',  $device], ['ADAS_VARIABLE_NAME', $var]];
	my $rs_device = $self->search_device($device_return_fields, $device_criteria);  #record set of device;
	my $value = 0;

	while(!$rs_device->EOF){
		my $device_id = $rs_device->Fields('ADAS_DEVICE')->{Value};
     		my $var_id = $rs_device->Fields('ADAS_VARIABLE')->{Value};
		my $LP_criteria =  [['ADAS_DEVICE', $device_id], ['ADAS_VARIABLE', $var_id], ['ADAS_TIME_GMT', $date_from, '>'], ['ADAS_TIME_GMT', $date_to, '<=']];
	my $rs = $self->{'MeterDataInterface'}->GetLoadProfile($LP_return_fields, $LP_criteria, 0) or die $hl->{'MeterDataInterface'}->LastError();
	while(!$rs->EOF){
		$value += $rs->Fields($LP_return_fields->[0])->{Value};

		warn $rs->Fields($LP_return_fields->[1])->{Value} . "meter value is" . $value  if DEBUG == 1;	
		$rs->MoveNext();
	}
	$rs_device->MoveNext();
}
	return $value;
}

=head2 get_single_LP
return a single value or a hash 
get_single_LP('ADAS_VAL_RAW', '海电_主表110', '+A', 'today');  # return a value
get_single_LP('ADAS_VAL_NORM', '主表%', 'yesterday');          # return a hash ref including all key/value fitting 主表%
=cut
sub get_single_LP{
	my $self = shift;
	my ($LP_return_fields, $device, $var, $date) = @_;
	    $LP_return_fields = [ $LP_return_fields ];
	my $parser = DateTime::Format::Natural->new(
	    			lang          => 'en',
	    			format     => 'yyyy/mm/dd',
	    			time_zone  => 'Asia/Taipei',
			);
    my $dt = $parser->parse_datetime($date);
       $dt->subtract(minutes => 1);
        $dt->set_time_zone( 'UTC' );
      my  $date_from = Variant(VT_DATE, 25569+$dt->epoch/86400); 
	my $date_to   = $self->convert_VT_DATE($date);
	my $device_return_fields = ['ADAS_DEVICE', 'ADAS_VARIABLE', 'ADAS_DEVICE_NAME','ADAS_VARIABLE_NAME'];
	my $device_criteria = [['ADAS_DEVICE_NAME',  $device], ['ADAS_VARIABLE_NAME', $var]];
	my $rs_device = $self->search_device($device_return_fields, $device_criteria);  #record set of device;
	my $LP_values ={};

	while(!$rs_device->EOF){
		my $device_id = $rs_device->Fields('ADAS_DEVICE')->{Value};
     		my $var_id = $rs_device->Fields('ADAS_VARIABLE')->{Value};
		my $LP_criteria =  [['ADAS_DEVICE', $device_id], ['ADAS_VARIABLE', $var_id], ['ADAS_TIME_GMT', $date_from, '>'], ['ADAS_TIME_GMT', $date_to, '<=']];
	my $rs = $self->{'MeterDataInterface'}->GetLoadProfile($LP_return_fields, $LP_criteria, 0) or die $hl->{'MeterDataInterface'}->LastError();
	while(!$rs->EOF){
		my $name = $rs_device->Fields('ADAS_DEVICE_NAME') . "_" . $rs_device->Fields('ADAS_VARIABLE_NAME');
		$LP_values->{$name} = $rs->Fields($LP_return_fields->[0])->{Value}; 	
		$rs->MoveNext();
	}
	$rs_device->MoveNext();
}
  return values(%LP_values) if keys(%LP_values) <=1;
  return \%LP_values;

}
=head2 search_device
return a recordset of device info fitting criteria 


e.g.:
while(!$rs->EOF){
...
$rs->MoveNext();
}
=cut

sub search_device{
    my $self = shift;
    my ($return_fields, $criteria) = @_;
    my $rs   = $self->{'MeterDataInterface'}->FindVariable($return_fields, $criteria, $Empty, 0) or die $self->{'MeterDataInterface'}->LastError();

    return $rs;
}





=head2  convert_VT_DATE 
C3000 utils
pass to a DateTime obj and return a VT_DATE variable.
=cut

sub convert_VT_DATE {
	shift;
use constant EPOCH       => 25569;
use constant SEC_PER_DAY => 86400;
    my $parser = DateTime::Format::Natural->new(
	    			lang          => 'en',
	    			format     => 'yyyy/mm/dd',
	    			time_zone  => 'Asia/Taipei',
			);
    my $date_string = shift;
    my $dt = $parser->parse_datetime($date_string);
    $dt->set_time_zone( 'UTC' );
    return Variant(VT_DATE, EPOCH+$dt->epoch/SEC_PER_DAY); 
}

=head1 AUTHOR

Andy Xiao, C<< <xyf.gmail.com> >>

=head1 BUGS

Please report any bugs or feature requests to C<bug-c3000 at rt.cpan.org>, or through
the web interface at L<http://rt.cpan.org/NoAuth/ReportBug.html?Queue=C3000>.  I will be notified, and then you'll
automatically be notified of progress on your bug as I make changes.




=head1 SUPPORT

You can find documentation for this module with the perldoc command.

    perldoc C3000


You can also look for information at:

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

1; # End of C3000
