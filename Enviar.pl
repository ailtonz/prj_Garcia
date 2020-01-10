#!c:/Perl/perl.exe -w

my $Lista = 'Arquivos.txt';
my $Cliente = 'Garcia | '.DataHora();
my $Pacote = DataHora().".zip";

CompactarArquivos($Lista,$Pacote);
EnviarDados('smtp.dominio.com.br','contato@dominio.com.br',$Cliente,$Pacote);


sub EnviarDados{
# SMTP, E-MAIL, TITULO, ARQUIVO

	use MIME::Lite;
	
	@ARQV = $_[3];

	# Configuração do E-mail.
	my $SMTP_SERVER = $_[0];
	my $DEFAULT_SENDER = $_[1];
	my $DEFAULT_RECIPIENT = $_[1];
	MIME::Lite->send('smtp', $SMTP_SERVER, Timeout=>60);

	my (%o, $msg);
	  
	$o{f} ||= $DEFAULT_SENDER;
	$o{t} ||= $DEFAULT_RECIPIENT;
	$o{s} ||= $_[2];

	# construct and send email
	$msg = new MIME::Lite(
	    From => $o{f},
	    To   => $o{t},
	    Subject => $o{s},
	    Data => "Hi",
	    Type => "multipart/mixed",
	);

	# Anexar  arquivo p/ envio.
	while (@ARQV) {
	  $msg->attach('Type' => 'application/octet-stream',
	               'Encoding' => 'base64',
	               'Path' => shift @ARQV);
	}

	# Enviar e-mail
	$msg->send();
	}

sub CompactarArquivos{
# ARQVS, NomeDoPacote
	# Create a Zip file
	use Getopt::Std;
	use Archive::Zip qw( :ERROR_CODES :CONSTANTS );
	my $zip = Archive::Zip->new();

	open(my $file, q{<}, $_[0]) or die "Can't open file $!\n";
	foreach my $arquivo ( <$file> ) {
	   #Remove o último caractere apenas se for igual a $/ "Separador de regsitro" 
	   chomp($arquivo);
	   # Verifica se o arquivo existe  
	   if (-e $arquivo){
	      # Add a file from disk
	      my $file_member = $zip->addFile($arquivo);
	   }
	}

	# Save the Zip file
	unless ( $zip->writeToFileNamed($_[1]) == AZ_OK ) {
	   die 'write error';
	}
	
	}	
	
sub DataHora{
# Criar layout de data
	my($dd,$mm,$yy,$day,$hh,$nn) = (localtime)[3,4,5,6,2,1];
	my $today =  join '', map sprintf("%02d", $_),($yy%100,$mm+1,$dd,);
	my $hr = join '', map sprintf("%02d", $_),($hh,$nn);

	# "Vasio" = Data e hora | "1" Apenas Data
	my $data = $_[0] ? $today:$today.'_'.$hr;
		
	my $datahora = $data;
	
	return ($datahora);
	}
