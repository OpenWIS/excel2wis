#!/usr/bin/perl
use strict;
use Spreadsheet::Read;
use Data::Dumper;
#http://search.cpan.org/~hmbrand/Spreadsheet-Read-0.58/Read.pm

# my $workbook  = ReadData ('Metadata-guide-record.v1.4.xls',parser => "xls");
# my $workbook  = ReadData ('Metadata-guide-record.v1.6_FF.ods',parser => "ods");
my $workbook  = ReadData ('Metadata-guide-record.v1.6_FF.xls');
my %infosOBLIGATOIRES;
my  %colTOxpath;

# print Dumper($workbook);
# print $workbook->[0]{sheet} . "\n";

#on commence par initialiser les données obligatoires
InitialiseDonneesObligatoires();

#on vérifie que le fichier excel est bien conforme à ce qui est attendu =les données obligatoires
my $onContinue=1;
my $MessageVerifieInfosObligatoires="";
if (! VerifieInfosObligatoires(\$MessageVerifieInfosObligatoires))
	{
	print $MessageVerifieInfosObligatoires."\n";
	$onContinue=0;
	}
else
	{
	print "On continue\n";
	}
#on remplit une table de correspdonce entre la colone de MDFields et les xpath
if ($onContinue)
	{
	coloneVersXpath();
	}
	
################################################################################
###################      FONCTIONS						################################
################################################################################
sub InitialiseDonneesObligatoires
	#procedure qui initialise à vide le hash des infos obligatoires: les noms des 4 onglets, les colones contenant les xpath
	{
	%infosOBLIGATOIRES=	(
			#noms/indices des onglets
		"tab_MDFields"				=> 	1,
		"tab_Help"					=> 	2,
		"tab_MDgeneric"				=> 	3,
		"tab_Thesaurus"				=> 	4,
			#colones renfermant des xpath
		"colone_xpath_MDgeneric"		=> 	"D",
		"colone_xpath_Help"			=> 	"H",
			#colone de Help.Section
		"colone_ids_sections"			=> 	"B",
			#1ère colone de MD_Fields contenant des données
		"1ereColoneRenseignee_MDFields"	=> 	"B",
			#1ère ligne de Help contenant des données
		"1ereLigneRenseignee_Help"		=> 	4,
			#1ère ligne de MD generic contenant des données
		"1ereLigneRenseignee_MDgeneric"	=> 	3,
						);
	}

################################################################################
# fonction qui vérifie que le fichier excel ressemble à ce qu'on attend
sub VerifieInfosObligatoires
	{
	my ($RefMonMessage)=@_;
	my $InfosObligatoiresCorrect=1;
	
		#verif indices des onglets
	$InfosObligatoiresCorrect=0 unless ($workbook->[0]{sheet}{"MD Fields"} eq "1");
	$InfosObligatoiresCorrect=0 unless ($workbook->[0]{sheet}{"Help"} eq "2" );
	$InfosObligatoiresCorrect=0 if ($workbook->[0]{sheet}{"MD generic"} ne "3");
	$InfosObligatoiresCorrect=0 if ($workbook->[0]{sheet}{"MD Thesaurus"} ne "4");
	$$RefMonMessage="Les onglets ne sont pas dans l ordre attendu.\n" unless $InfosObligatoiresCorrect ;

		#verif cases xpath
	my $ligneXpath=$infosOBLIGATOIRES{"1ereLigneRenseignee_Help"}-2;
	my $caseXpath=$infosOBLIGATOIRES{"colone_xpath_Help"}.$ligneXpath;
	if ($workbook->[$infosOBLIGATOIRES{tab_Help}]{$caseXpath}!~/XPATH/i)
		{
		$InfosObligatoiresCorrect=0;
		$$RefMonMessage=$$RefMonMessage."La case ".$workbook->[$infosOBLIGATOIRES{tab_Help}]{label}."\.$caseXpath devrait contenir XPATH\n";
		}
	$ligneXpath=$infosOBLIGATOIRES{"1ereLigneRenseignee_MDgeneric"}-1;
	$caseXpath=$infosOBLIGATOIRES{"colone_xpath_MDgeneric"}.$ligneXpath;
	if ($workbook->[$infosOBLIGATOIRES{tab_MDgeneric}]{$caseXpath}!~/XPATH/i)
		{
		$InfosObligatoiresCorrect=0;
		$$RefMonMessage=$$RefMonMessage."La case ".$workbook->[$infosOBLIGATOIRES{tab_MDgeneric}]{label}."\.$caseXpath devrait contenir XPATH\n";
		}
	
	return $InfosObligatoiresCorrect;
	}
	
################################################################################
# fonction qui remplit le hash %colTOxpath permettant de mapper la colone de MD Fields vers le xpath décrit dans Help
# exemple : D	/gmd:MD_Metadata/gmd:fileIdentifier
sub coloneVersXpath
	{
		#on commence par remplir le hash de la ligne 4 de l'onglet 'MD Fields' %ligne4 (clef: contenu de la colone que l'on retrouve dans l'onglet help, valeur=Colone)
	my $cestNonVide=1;
	my $idCellule;
	my $iCol=$infosOBLIGATOIRES{"1ereColoneRenseignee_MDFields"};
	my %ligne4;
	my $contenuCelluleEnCours;
	
	while ($cestNonVide)
		{
		$idCellule=$iCol.$infosOBLIGATOIRES{"1ereLigneRenseignee_Help"};
		$contenuCelluleEnCours=$workbook->[$infosOBLIGATOIRES{tab_MDFields}]{$idCellule};
		if ($contenuCelluleEnCours)
			{
			$ligne4{$contenuCelluleEnCours}=$iCol;
			# print $iCol."\t".$contenuCelluleEnCours."\n";
			$iCol++;
			}
		else
			{
			$cestNonVide=0;
			}
		}
		#on va remplir le hash %colTOxpath (clef: numéro de colone dans l'onglet 'MD Fields' , valeur : valeur du xpath lu dans l'onglet 'Help'
	my $iLigne=$infosOBLIGATOIRES{"1ereLigneRenseignee_Help"};
	my $ya1xpath=1;
	my $idCelluleSection;
	my $section;
	while ($ya1xpath)
		{
		$idCellule=$infosOBLIGATOIRES{colone_xpath_Help}.$iLigne;
		$idCelluleSection=$infosOBLIGATOIRES{colone_ids_sections}.$iLigne;
		$contenuCelluleEnCours=$workbook->[$infosOBLIGATOIRES{tab_Help}]{$idCellule};
		$section=$workbook->[$infosOBLIGATOIRES{tab_Help}]{$idCelluleSection};
		if ($contenuCelluleEnCours)
			{
			if ($contenuCelluleEnCours=~/^\/\/*/)
				{
				$colTOxpath{$ligne4{$section}}=$contenuCelluleEnCours;
				print $ligne4{$section}."\t".$contenuCelluleEnCours."\n";
				$iLigne++;	
				}
			}
		else
			{
			$ya1xpath=0;
			}
		}
	}
