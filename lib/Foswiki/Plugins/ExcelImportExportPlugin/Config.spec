# ---+ Extensions
# ---++ ExcelImportExportPlugin
# This is the configuration used by the <b>ExcelImportExportPlugin</b>.

# **PERL EXPERT**
# This setting is required to enable executing the excel2topics service 
$Foswiki::cfg{SwitchBoard}{excel2topics} = {
  'function' => 'excel2topics',
  'context' => {
    'excel2topics' => 1,
  },
  'package' => 'Foswiki::Plugins::ExcelImportExportPlugin::Import'
};

# **PERL EXPERT**
# This setting is required to enable executing the excel2topics service 
$Foswiki::cfg{SwitchBoard}{topics2excel} = {
  'function' => 'topics2excel',
  'context' => {
    'topics2excel' => 1,
  },
  'package' => 'Foswiki::Plugins::ExcelImportExportPlugin::Export'
};

# **PERL EXPERT**
# This setting is required to enable executing the excel2topics service 
$Foswiki::cfg{SwitchBoard}{table2excel} = {
  'function' => 'table2excel',
  'context' => {
    'table2excel' => 1,
  },
  'package' => 'Foswiki::Plugins::ExcelImportExportPlugin::Export'
};

# **PERL EXPERT**
# This setting is required to enable executing the excel2topics service 
$Foswiki::cfg{SwitchBoard}{uploadexcel2table} = {
  'function' => 'uploadexcel2table',
  'context' => {
    'uploadexcel2table' => 1,
  },
  'package' => 'Foswiki::Plugins::ExcelImportExportPlugin::Import'
};

1;
