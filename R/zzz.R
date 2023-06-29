# script: scr_zzz
# author: Serkan Korkmaz
# date: 2023-06-21
# objective: A set of functions to
# be executed upon loading
# script start; ####

# pkg load
.onAttach <- function(libname, pkgname) {
  
  # TODO: Fix later
  
  packageStartupMessage(
    rlang::inform(
      msg=  cli::cli_rule(
        left = cli::col_br_magenta(
          'workbookR'
        ),
        right = cli::col_magenta(cli::style_bold(
          'Version: '
        ),
        '1.0'
        )
      )
      , class = "packageStartupMessage")
  )
  
  
  
  
  
}



# end of script; ####