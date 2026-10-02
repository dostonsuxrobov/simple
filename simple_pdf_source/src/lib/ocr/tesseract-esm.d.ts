// tesseract.js ships its ESM build without typings; it is the same API as the
// package entry (a default export holding createWorker, OEM, PSM, ...).
declare module 'tesseract.js/dist/tesseract.esm.min.js' {
  import Tesseract from 'tesseract.js'

  export default Tesseract
}
