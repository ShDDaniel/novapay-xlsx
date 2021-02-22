const fs = require('fs');
const path = require('path');
const { Transform } = require('stream');
const archiver = require('archiver');
const converter = require('xml-js');

const UTF16_A = 65;
const ALPHABET_TOTAL_CHAR_COUNT = 26;
const CELL_TYPES = {
	BOOL: 'b',
	NUMBER: 'n',
	INLINE_STR: 'inlineStr'
};

const initChunk = fs.readFileSync(path.join(__dirname, 'src/xml/initChunk.xml'));
const finalChunk = fs.readFileSync(path.join(__dirname, 'src/xml/finalChunk.xml'));

/**
 * 	Returns a column's literal index as an array of chars,
 *	e.g.: 0 -> [A], 27 -> [A, A], 28 -> [A, B] ...
 * 	@param {number} numIndex Index of a config item
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

const isNumber = (val) => {
	return typeof val === 'number' || typeof parseInt(val) === 'number' || typeof parseFloat(val) === 'number';
};

const isBool = (val) => {
	return typeof val === 'boolean' || val === 'true' || val === 'false';
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
					let val = c.formatter ? c.formatter(row[c.key], row) : row[c.key];

					if (isBool(val)) {
						return {
							_attributes: { r, t: CELL_TYPES.BOOL },
							v: { _text: [true, 'true'].includes(val) ? 1 : 0 }
						};
					}
					if (isNumber(val)) {
						return {
							_attributes: { r, t: CELL_TYPES.NUMBER },
							v: { _text: val }
						};
					}
					return {
						_attributes: { r, t: CELL_TYPES.INLINE_STR },
						is: {
							t: {
								_text: val || ''
							}
						}
					};
				})
			}
		};
		this.push(Buffer.from(converter.json2xml(json, { compact: true })));
		this.row++;
	}
	_transform(chunk, encoding, callback) {
		if (this.row === 1) {
			this.push(initChunk);
			this.addHeader();
			this.addRow(this.options.chunkRowKey ? chunk[this.options.chunkRowKey] : chunk);
			callback();
		} else {
			this.addRow(this.options.chunkRowKey ? chunk[this.options.chunkRowKey] : chunk);
			callback();
		}
	}
	_flush(callback) {
		this.push(finalChunk);
		callback();
	}
}

/**
 * @param {(ReadableStream|PassThrough|TransformStream)} source
 *
 * @param {Object[]} config
 * @param {string} config[].key
 * @param {string} config[].label
 * @param {function} config[].formatter
 *
 * @param {Object} [options]
 * @param {boolean} options.debugMemUsage
 * @param {string} options.chunkRowKey
 */
const xlsx = (source, config, options = {}) => {
	if (!Array.isArray(config)) {
		throw new Error('config should be an array of objects');
	}
	config.forEach((c) => {
		if (!c || !c.key || !c.label) {
			let err =
				`a config item\n ${JSON.stringify(c)}\n is either missing` +
				`'key' or 'label' properties or they have invalid values.`;
			throw new Error(err);
		}
		if (c.formatter && typeof c.formatter !== 'function') {
			throw new Error(`'formatter' in \n ${JSON.stringify(c)}\n is not a function.`);
		}
	});

	let worksheetPipe = new WorksheetWriter({ config, options });

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
	archive.finalize().catch((err) => {
		archive.emit('error', err);
	});

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
