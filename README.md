xlsx stream writer

example usage:

xlsx(source, config, { ...options }).pipe(require('fs').createWriteStream('output.xlsx'))