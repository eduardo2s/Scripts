library(RCurl)
library(curl)
url <- "ftp://link"
userpwd <- "user:pass"
filenames <- getURL(url, userpwd = userpwd, ftp.use.epsv = FALSE,dirlistonly = TRUE)
files <- unlist(strsplit(filenames, '\r\n'))
dest <- "C:pasta/onde/salvar/arquivo"
for (filename in files) {
  download.file(paste(url, filename, sep = ""), paste(dest, "/", filename, sep = ""))
}

# test
download.file(paste(url, files[3], sep = ""), paste(dest, "/", files[3], sep = ""), mode = 'wb')
