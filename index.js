const fs = require('fs');
const path = require('path');
const { Transform } = require('stream');
const archiver = require('archiver');
const converter = require('xml-js');

const UTF16_A = 65;
const ALPHABET_TOTAL_CHAR_COUNT = 26;

/*
	Get column literal index by numeric index.
	e.g.: 0 -> A, 27 -> AA, 28 -> AB ...
*/
const getCharIndexByNumIndex = (numIndex) => {
	let divRes = Math.floor(numIndex / ALPHABET_TOTAL_CHAR_COUNT);
	let divRem = Math.floor(numIndex % ALPHABET_TOTAL_CHAR_COUNT);
	if (!divRes && !divRem) {
		return ['A'];
	}
	if (divRes) {
		// prettier-ignore
		let char1 = divRes > ALPHABET_TOTAL_CHAR_COUNT
			? getCharIndexByNumIndex(divRes - 1)
			: [String.fromCharCode(UTF16_A + divRes - 1)];
		let char2 = String.fromCharCode(UTF16_A + divRem);
		return [...char1, char2];
	}
	if (!divRes && divRem) {
		let char = String.fromCharCode(UTF16_A + divRem);
		return [char];
	}
};

class WorksheetWriter extends Transform {
	constructor({ config, options }) {
		super({ objectMode: true });
		this.row = 1;

		this.config = config;
		this.options = options;
	}
	generateCellIndex(numIndex) {
		let chars = getCharIndexByNumIndex(numIndex);
		return `${chars.join('')}${this.row}`;
	}
	addHeader() {
		let json = {
			row: {
				_attributes: {
					r: this.row,
					customFormat: false,
					ht: 12.8,
					hidden: false,
					customHeight: false,
					outlineLevel: 0,
					collapsed: false
				},
				c: this.config.map(({ label }, index) => {
					let r = this.generateCellIndex(index);
					return {
						_attributes: { r, t: 'inlineStr' },
						is: {
							t: {
								_text: label
							}
						}
					};
				})
			}
		};
		this.push(Buffer.from(converter.json2xml(json, { compact: true })));
		this.row++;
	}
	addRow(row) {
		let json = {
			row: {
				_attributes: {
					r: this.row,
					customFormat: false,
					ht: 12.8,
					hidden: false,
					customHeight: false,
					outlineLevel: 0,
					collapsed: false
				},
				c: this.config.map((c, index) => {
					let r = this.generateCellIndex(index);
					return {
						_attributes: { r, t: 'inlineStr' },
						is: {
							t: {
								_text: (c.formatter ? c.formatter(row[c.key], row) : row[c.key]) || ''
							}
						}
					};
				})
			}
		};
		this.push(Buffer.from(converter.json2xml(json, { compact: true })));
		this.row++;
	}
	finalizeSheet() {
		this.push(fs.readFileSync(path.join(__dirname, 'src/xml/finalChunk.xml')));
	}
	initSheet() {
		this.push(fs.readFileSync(path.join(__dirname, 'src/xml/initChunk.xml')));
	}
	_transform(chunk, encoding, callback) {
		if (this.row === 1) {
			this.initSheet();
			this.addHeader();
			this.addRow(this.options.chunkRowKey ? chunk[this.options.chunkRowKey] : chunk);
			callback();
		} else {
			this.addRow(this.options.chunkRowKey ? chunk[this.options.chunkRowKey] : chunk);
			callback();
		}
	}
}

const xlsx = (source, config, options = {}) => {
	if (!Array.isArray(config)) {
		throw new Error('config should be an array of objects');
	}

	let worksheetPipe = new WorksheetWriter({ config, options });
	source.on('end', () => {
		worksheetPipe.finalizeSheet();
	});

	const archive = archiver('zip', { zlib: { level: 9 } });
	archive.directory(path.join(__dirname, 'src/xml/_rels'), '_rels');
	archive.directory(path.join(__dirname, 'src/xml/docProps'), 'docProps');
	archive.directory(path.join(__dirname, 'src/xml/xl'), 'xl');
	// prettier-ignore
	archive.append(
		fs.createReadStream(path.join(__dirname, '/src/xml/[Content_Types].xml')),
		{ name: '[Content_Types].xml' }
	);
	archive.append(source.pipe(worksheetPipe), { name: 'xl/worksheets/sheet1.xml' });
	archive.finalize();

	archive.on('end', () => {
		if (options.debugMemUsage) {
			let used = process.memoryUsage();
			let keys = Object.keys(used);
			keys.forEach((key) => {
				// eslint-disable-next-line
				console.log(`${key} ${Math.round(used[key] / 1024 / 1024)} MB`);
			});
		}
	});
	archive.on('warning', (err) => {
		archive.emit('error', err);
	});
	worksheetPipe.on('error', (err) => {
		archive.emit('error', err);
	});
	source.on('error', (err) => {
		archive.emit('error', err);
	});
	return archive;
};

module.exports = xlsx;
