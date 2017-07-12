#If corrections are made or Mike suggests you update your version of the script,
#you can just run these commands.


remove.packages(pkgs = "bio.portsampling")
require(devtools)
install_github('Maritimes/bio.portsampling')
require(bio.portsampling)
makeHailInRpt()
#makeHailInRpt(thePath = "~/doot")
