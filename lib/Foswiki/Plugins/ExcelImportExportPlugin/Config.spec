# ---+ Extensions
# ---++ ExcelImportExportPlugin
# This is the configuration used by the <b>ExcelImportExportPlugin</b>.

# **PERL EXPERT**
# This setting is required to enable executing the excel2topics service 
$Foswiki::cfg{SwitchBoard}{excel2topics} = {
  'function' => 'excel2topics',
  'context' => {
    'view' => 1,
  },
  'package' => 'Foswiki::Plugins::ExcelImportExportPlugin::Import'
};

# **PERL EXPERT**
# This setting is required to enable executing the excel2topics service 
$Foswiki::cfg{SwitchBoard}{topics2excel} = {
  'function' => 'topics2excel',
  'context' => {
    'view' => 1,
  },
  'package' => 'Foswiki::Plugins::ExcelImportExportPlugin::Export'
};

1;
