const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  PageNumber, PageBreak, LevelFormat, Footer, TabStopType, VerticalAlign
} = require('docx');
const fs = require('fs');

// ── spacing constants ───────────────────────────────────────────────────────
const LINE_15 = { line: 360, lineRule: 'auto' }; // 1.5 line spacing (req. G)

// ── text helpers ────────────────────────────────────────────────────────────
const t  = (text, extra = {}) => new TextRun({ text, font: 'Arial', size: 24, ...extra });
const tb = (text, extra = {}) => new TextRun({ text, font: 'Arial', size: 24, bold: true, ...extra });
const ti = (text) => new TextRun({ text, font: 'Arial', size: 24, italics: true });

// creatividad in bold - used every analytical reference (req. A)
const CLAVE = () => tb('creatividad');

// Body paragraph with 1.5 spacing
function bp(runs, before = 80, after = 120) {
  if (typeof runs === 'string') runs = [t(runs)];
  return new Paragraph({
    children: runs,
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before, after, ...LINE_15 },
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, font: 'Arial', size: 28, bold: true, color: '1A3C5E' })],
    spacing: { before: 320, after: 120 },
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, font: 'Arial', size: 26, bold: true, italics: true, color: '2C5F8A' })],
    spacing: { before: 200, after: 80 },
  });
}

const thin = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const allB = { top: thin, bottom: thin, left: thin, right: thin };
const noB  = { top:  { style: BorderStyle.NONE, size:0, color:'FFFFFF' },
               bottom:{ style: BorderStyle.NONE, size:0, color:'FFFFFF' },
               left:  { style: BorderStyle.NONE, size:0, color:'FFFFFF' },
               right: { style: BorderStyle.NONE, size:0, color:'FFFFFF' } };

function coverRow(label, value, fill = 'FFFFFF') {
  const mkCell = (txt, w, bold_ = false, shade = fill) => new TableCell({
    borders: allB,
    width: { size: w, type: WidthType.DXA },
    shading: { fill: shade, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      children: [new TextRun({ text: txt, font: 'Arial', size: 22, bold: bold_ })],
      spacing: { before: 0, after: 0 },
    })],
  });
  return new TableRow({ children: [mkCell(label, 2400, true, 'E8F0FB'), mkCell(value, 6626)] });
}

// Aaker table helper
function aRow(dim, val, i) {
  const fill = i % 2 === 0 ? 'EAF1FB' : 'FFFFFF';
  return new TableRow({
    children: [
      new TableCell({
        borders: allB, width: { size: 2200, type: WidthType.DXA },
        shading: { fill, type: ShadingType.CLEAR },
        margins: { top: 70, bottom: 70, left: 100, right: 100 },
        children: [new Paragraph({ children: [tb(dim, { size: 22 })], spacing: { before: 0, after: 0 } })],
      }),
      new TableCell({
        borders: allB, width: { size: 6826, type: WidthType.DXA },
        shading: { fill, type: ShadingType.CLEAR },
        margins: { top: 70, bottom: 70, left: 100, right: 100 },
        children: [new Paragraph({ children: [t(val, { size: 22 })], alignment: AlignmentType.JUSTIFIED, spacing: { before: 0, after: 0 } })],
      }),
    ],
  });
}

function aHeader(c1, c2) {
  return new TableRow({
    children: [c1, c2].map((label, i) => new TableCell({
      borders: allB,
      width: { size: i === 0 ? 2200 : 6826, type: WidthType.DXA },
      shading: { fill: '1A3C5E', type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 100, right: 100 },
      children: [new Paragraph({
        children: [new TextRun({ text: label, font: 'Arial', size: 22, bold: true, color: 'FFFFFF' })],
        alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 },
      })],
    })),
  });
}

// ── DOCUMENT ─────────────────────────────────────────────────────────────────
const doc = new Document({
  styles: {
    default: { document: { run: { font: 'Arial', size: 24 } } },
    paragraphStyles: [
      {
        id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 28, bold: true, font: 'Arial', color: '1A3C5E' },
        paragraph: { spacing: { before: 320, after: 120 }, outlineLevel: 0 },
      },
      {
        id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 26, bold: true, italics: true, font: 'Arial', color: '2C5F8A' },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 1 },
      },
    ],
  },

  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
    },

    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            t('Fenty Beauty \u2022 Concepto clave: Creatividad   \u2022   P\u00e1gina ', { size: 20 }),
            new TextRun({ font: 'Arial', size: 20, children: [PageNumber.CURRENT] }),
          ],
        })],
      }),
    },

    children: [

      // ════════════════════════════════════════════════════════════════════
      // PORTADA (anonimizada — req. G)
      // ════════════════════════════════════════════════════════════════════
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 720, after: 80 },
        children: [new TextRun({ text: 'COLEGIO ALTAIR', font: 'Arial', size: 36, bold: true, color: '1A3C5E' })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 80 },
        children: [ti('An Inspired School')],
      }),
      new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '1A3C5E', space: 1 } },
        spacing: { before: 0, after: 360 }, children: [],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 60 },
        children: [new TextRun({ text: 'PROYECTO DE INVESTIGACI\u00d3N EMPRESARIAL', font: 'Arial', size: 30, bold: true })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 360 },
        children: [ti('IB Diploma Programme \u2014 Gesti\u00f3n Empresarial (Nivel Medio)')],
      }),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2400, 6626],
        rows: [
          coverRow('Asignatura', 'Gesti\u00f3n Empresarial \u2014 Nivel Medio (NM)', 'E8F0FB'),
          coverRow('Concepto clave', 'Creatividad', 'FFFFFF'),
          coverRow('Organizaci\u00f3n', 'Fenty Beauty (LVMH)', 'E8F0FB'),
          coverRow('Pregunta de investigaci\u00f3n',
            '\u00bfDe qu\u00e9 manera la creatividad en la estrategia de marketing inclusivo impact\u00f3 en el posicionamiento de Fenty Beauty en la industria cosm\u00e9tica?',
            'FFFFFF'),
          coverRow('Convocatoria', 'Mayo 2026', 'E8F0FB'),
          coverRow('N\u00famero de candidato/a', '______________________', 'FFFFFF'),
          coverRow('C\u00f3mputo de palabras', '1\u202f795 palabras', 'E8F0FB'),
        ],
      }),

      new Paragraph({ children: [new PageBreak()], spacing: { before: 0, after: 0 } }),

      // ════════════════════════════════════════════════════════════════════
      // DECLARACI\u00d3N DE AUTENTICIDAD (req. G)
      // ════════════════════════════════════════════════════════════════════
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 720, after: 240 },
        children: [new TextRun({ text: 'DECLARACI\u00d3N DE AUTENTICIDAD', font: 'Arial', size: 26, bold: true })],
      }),
      bp([
        t('Declaro que este proyecto de investigaci\u00f3n empresarial es mi trabajo original y que en su redacci\u00f3n no he infringido ning\u00fan reglamento del IB relacionado con la conducta deshonesta. Todas las fuentes utilizadas han sido debidamente citadas y referenciadas en la bibliograf\u00eda. He reconocido el uso de inteligencia artificial u otras herramientas digitales de acuerdo con las pol\u00edticas vigentes del IB y del colegio.'),
      ], 40, 80),
      bp([
        t('Firma del alumno/a: ___________________________\u2003Fecha: _________________'),
      ], 240, 40),
      bp([
        t('Firma del profesor/a: ___________________________\u2003Fecha: _________________'),
      ], 40, 40),

      new Paragraph({ children: [new PageBreak()], spacing: { before: 0, after: 0 } }),

      // ════════════════════════════════════════════════════════════════════
      // \u00cdNDICE
      // ════════════════════════════════════════════════════════════════════
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 240, after: 240 },
        children: [new TextRun({ text: '\u00cdNDICE DE CONTENIDOS', font: 'Arial', size: 26, bold: true })],
      }),
      ...[
        ['1.\u2002Introducci\u00f3n', '4'],
        ['\u2003\u20021.1\u2002Contexto empresarial', '4'],
        ['\u2003\u20021.2\u2002Problema identificado y metodolog\u00eda', '4'],
        ['2.\u2002Cuerpo principal', '5'],
        ['\u2003\u20022.1\u2002La creatividad en el Mix de Marketing (4Ps)', '5'],
        ['\u2003\u20022.2\u2002Impacto competitivo: Diferenciaci\u00f3n de Porter y Mapa de Posicionamiento', '6'],
        ['\u2003\u20022.3\u2002Evaluaci\u00f3n cr\u00edtica: \u00bfes duradera la ventaja?', '6'],
        ['\u2003\u20022.4\u2002La ventaja intangible: Marco de Equidad de Marca de Aaker', '7'],
        ['3.\u2002Conclusi\u00f3n', '8'],
        ['Bibliograf\u00eda', '9'],
        ['Gu\u00eda de documentos de apoyo', '10'],
      ].map(([label, page]) => new Paragraph({
        alignment: AlignmentType.LEFT,
        tabStops: [{ type: TabStopType.RIGHT, position: 9026 }],
        spacing: { before: 40, after: 40, ...LINE_15 },
        children: [t(label, { size: 22 }), t('\t' + page, { size: 22 })],
      })),

      new Paragraph({ children: [new PageBreak()], spacing: { before: 0, after: 0 } }),

      // ════════════════════════════════════════════════════════════════════
      // 1. INTRODUCCI\u00d3N
      // ════════════════════════════════════════════════════════════════════
      h1('1. Introducci\u00f3n'),
      h2('1.1 Contexto empresarial'),

      bp([
        t('Fenty Beauty fue fundada en septiembre de 2017 por la artista Robyn Rihanna Fenty en asociaci\u00f3n con LVMH Mo\u00ebt Hennessy Louis Vuitton. Desde su concepci\u00f3n, la marca declar\u00f3 expl\u00edcitamente que su prop\u00f3sito era \u201ccelebrar la belleza de cada persona sin importar su tono de piel\u201d '),
        tb('(DA 3)', { size: 24 }),
        t('. Esta filosof\u00eda se materializ\u00f3 en el lanzamiento simult\u00e1neo en 17 pa\u00edses con 40 tonos de base \u2014posteriormente ampliados a 50\u2014, respondiendo a un vac\u00edo hist\u00f3rico en el mercado cosm\u00e9tico: competidores como MAC, L\u2019Or\u00e9al y Make Up For Ever concentraban aproximadamente el 77\u00a0% de sus gamas en pieles claras y medias '),
        tb('(DA 1)', { size: 24 }),
        t('. En 2023, Fenty Beauty gener\u00f3 ingresos globales de $602,4 millones USD, posicion\u00e1ndose como la marca de belleza de celebridad con mayores ingresos a nivel mundial '),
        tb('(DA 2)', { size: 24 }),
        t('.'),
      ]),

      h2('1.2 Problema identificado y metodolog\u00eda'),

      bp([
        t('En un mercado cosm\u00e9tico global proyectado a $580\u00a0000 millones para 2027, donde la saturaci\u00f3n convierte la diferenciaci\u00f3n en un reto estructural '),
        tb('(DA 5)', { size: 24 }),
        t(', esta investigaci\u00f3n examina la siguiente pregunta: '),
        ti('\u00bfde qu\u00e9 manera la '),
        new TextRun({ text: 'creatividad', font: 'Arial', size: 24, italics: true, bold: true }),
        ti(' en la estrategia de marketing inclusivo impact\u00f3 en el posicionamiento de Fenty Beauty en la industria cosm\u00e9tica?'),
        t(' El concepto clave empleado es '),
        CLAVE(),
        t(', definida como el proceso de generar nuevas ideas y considerar perspectivas existentes desde enfoques innovadores, con capacidad para reconocer el valor de dichas ideas en la elaboraci\u00f3n de respuestas a problemas (IB Gu\u00eda de Gesti\u00f3n Empresarial, 2022). Se argumenta que la '),
        CLAVE(),
        t(' fue el motor estrat\u00e9gico que permiti\u00f3 a Fenty Beauty identificar y ocupar un espacio de mercado desatendido, construyendo una propuesta de valor \u00fanica cuya durabilidad merece evaluaci\u00f3n cr\u00edtica.'),
      ]),

      bp([
        t('La investigaci\u00f3n se sustenta en cinco documentos de apoyo secundarios y uno interno: el sitio oficial de Fenty Beauty '),
        tb('(DA 3)', { size: 24 }),
        t('; un art\u00edculo acad\u00e9mico sobre posicionamiento estrat\u00e9gico e inclusividad '),
        tb('(DA 1)', { size: 24 }),
        t('; datos cuantitativos de ingresos de marcas de celebridad '),
        tb('(DA 2)', { size: 24 }),
        t('; y dos informes de McKinsey & Company sobre el consumidor inclusivo y el mercado cosm\u00e9tico global '),
        tb('(DA 4, DA 5)', { size: 24 }),
        t('. Las herramientas de Gesti\u00f3n Empresarial aplicadas son: Mix de Marketing (4Ps), Estrategia de Diferenciaci\u00f3n de Porter, Mapa de Posicionamiento y Marco de Equidad de Marca de Aaker.'),
      ]),

      new Paragraph({ children: [new PageBreak()], spacing: { before: 0, after: 0 } }),

      // ════════════════════════════════════════════════════════════════════
      // 2. CUERPO PRINCIPAL
      // ════════════════════════════════════════════════════════════════════
      h1('2. Cuerpo principal'),
      h2('2.1 La creatividad en el Mix de Marketing (4Ps)'),

      bp([
        t('El an\u00e1lisis del Mix de Marketing (Kotler & Armstrong, 2021) revela que la '),
        CLAVE(),
        t(' en Fenty Beauty no fue un recurso puntual sino una perspectiva estrat\u00e9gica transversal a las cuatro variables del modelo, lo que explica la coherencia y profundidad de su propuesta de valor.'),
      ]),

      bp([
        tb('Producto: '),
        t('La decisi\u00f3n m\u00e1s disruptiva fue el dise\u00f1o de la l\u00ednea de bases. Cuando la industria operaba bajo el supuesto impl\u00edcito de que 15-20 tonos eran suficientes, Fenty Beauty introdujo creativamente 40 tonos distribuidos de manera equitativa a lo largo del espectro crom\u00e1tico, cubriendo pieles muy oscuras y muy claras con igual representaci\u00f3n '),
        tb('(DA 1)', { size: 24 }),
        t('. La propia marca afirm\u00f3 en su portal oficial que esta decisi\u00f3n naci\u00f3 de la experiencia personal de Rihanna al no encontrar productos adecuados para su tono '),
        tb('(DA 3)', { size: 24 }),
        t(': la '),
        CLAVE(),
        t(' aqu\u00ed fue aplicada al proceso de desarrollo del producto, reencuadrando la pregunta de dise\u00f1o desde \u201c\u00bfqu\u00e9 tonos se venden m\u00e1s?\u201d hacia \u201c\u00bfqu\u00e9 tonos ha ignorado hist\u00f3ricamente la industria?\u201d. El reconocimiento de la revista '),
        ti('Time'),
        t(' como una de las 25 mejores invenciones de 2017 valid\u00f3 el car\u00e1cter innovador de la propuesta.'),
      ]),

      bp([
        tb('Promoci\u00f3n: '),
        t('Fenty Beauty adopt\u00f3 tres decisiones de comunicaci\u00f3n creativamente no convencionales. Primero, eligi\u00f3 micro-influencers con tonos de piel oscuros en lugar de las celebridades m\u00e1s visibles del segmento \u2014decisi\u00f3n contraria al modelo dominante de la industria\u2014. Segundo, cre\u00f3 la \u201cFenty Family\u201d, una comunidad que produjo contenido aut\u00e9ntico en redes sociales, en l\u00ednea con la preferencia documentada del consumidor inclusivo por recomendaciones de pares sobre publicidad tradicional '),
        tb('(DA 4)', { size: 24 }),
        t('. Tercero, en 2020 lanz\u00f3 la \u201cFenty Beauty House\u201d en Los \u00c1ngeles, donde influencers de TikTok generaron contenido org\u00e1nico con tasas de apertura de email superiores al 23,8\u00a0% del promedio de la industria '),
        tb('(DA 1)', { size: 24 }),
        t('.'),
      ]),

      bp([
        tb('Precio y Plaza: '),
        t('Con un rango mid-range de $22 a $40 por base, Fenty articul\u00f3 creativamente un posicionamiento de \u201clujo accesible\u201d que rompi\u00f3 con la dicotom\u00eda tradicional lujo-exclusividad versus masividad. El lanzamiento simult\u00e1neo en Sephora en 17 pa\u00edses \u2014en lugar del modelo convencional de expansi\u00f3n gradual\u2014 amplific\u00f3 el impacto global: una apuesta inusual para una marca reci\u00e9n fundada que demuestra la '),
        CLAVE(),
        t(' tambi\u00e9n en las decisiones de distribuci\u00f3n '),
        tb('(DA 3)', { size: 24 }),
        t('.'),
      ]),

      h2('2.2 Impacto competitivo: Diferenciaci\u00f3n de Porter y Mapa de Posicionamiento'),

      bp([
        t('La Estrategia de Diferenciaci\u00f3n de Porter (1985) postula que la ventaja competitiva surge cuando el consumidor percibe un valor superior que justifica la elecci\u00f3n de la marca. La '),
        CLAVE(),
        t(' de Fenty Beauty articu l\u00f3 esta diferenciaci\u00f3n en dos niveles complementarios.'),
      ]),

      bp([
        t('A nivel '),
        tb('funcional'),
        t(', la distribuci\u00f3n equitativa de tonos fue una ventaja cuantificable y visible: ning\u00fan competidor directo igualaba en 2017 la cobertura del espectro crom\u00e1tico de Fenty Beauty '),
        tb('(DA 1)', { size: 24 }),
        t('. Esta superioridad t\u00e9cnica, producto de la '),
        CLAVE(),
        t(' en el desarrollo de producto, gener\u00f3 una percepci\u00f3n de calidad e inclusividad que sus rivales no pod\u00edan igualar inmediatamente. A nivel '),
        tb('emocional'),
        t(', la '),
        CLAVE(),
        t(' oper\u00f3 como mecanismo de representaci\u00f3n identitaria. Seg\u00fan McKinsey, dos de cada tres consumidores estadounidenses declaran que sus valores sociales determinan sus decisiones de compra '),
        tb('(DA 4)', { size: 24 }),
        t('. Fenty no solo vendi\u00f3 maquillaje; construy\u00f3 creativamente una narrativa de representaci\u00f3n que gener\u00f3 lealtad emocional dif\u00edcilmente replicable mediante estrategias puramente funcionales.'),
      ]),

      bp([
        t('El '),
        tb('Mapa de Posicionamiento'),
        t(' (ejes: nivel de inclusividad / precio-calidad percibida) confirma este an\u00e1lisis. Maybelline ofre c\u00eda alta inclusividad a bajo precio; MAC, inclusividad media a alto precio; NARS, baja inclusividad a alto precio. Fenty Beauty identific\u00f3 y ocup\u00f3 creativamente el cuadrante vac\u00edo: m\u00e1xima inclusividad con precio mid-premium, imposible sin la '),
        CLAVE(),
        t(' que reencuadr\u00f3 el eje de competencia desde la representaci\u00f3n, no solo desde el precio o la calidad t\u00e9cnica.'),
      ]),

      h2('2.3 Evaluaci\u00f3n cr\u00edtica: \u00bfes duradera la ventaja?'),

      bp([
        t('Una evaluaci\u00f3n rigurosa exige cuestionar si la ventaja competitiva basada en la '),
        CLAVE(),
        t(' es duradera a largo plazo. El denominado \u201cFenty Effect\u201d desencaden\u00f3 una respuesta masiva de la industria: en menos de 18 meses tras el lanzamiento, m\u00faltiples marcas ampliaron sus gamas a m\u00e1s de 40 tonos '),
        tb('(DA 1)', { size: 24 }),
        t('. Esta r\u00e1pida imitaci\u00f3n demuestra que la '),
        CLAVE(),
        t(' expresada en caracter\u00edsticas tangibles del producto es imitable a corto plazo.'),
      ]),

      bp([
        t('Para 2024, Fenty Beauty compite directamente con Rare Beauty y Rhode en el mismo espacio de alta inclusividad '),
        tb('(DA 1)', { size: 24 }),
        t('. El art\u00edculo acad\u00e9mico de Kravchuk y Krupskyi '),
        tb('(DA 1)', { size: 24 }),
        t(' se\u00f1ala que la mera promesa de inclusividad ya no es suficiente para diferenciarse: las marcas deben integrarla aut\u00e9nticamente en toda su arquitectura de marca. Esto plantea un supuesto cr\u00edtico: \u00bfpuede Fenty Beauty sostener su '),
        CLAVE(),
        t(' diferenciadora ante competidores que han aprendido a replicar su modelo de inclusividad? Este interrogante permanece abierto y requerir\u00eda investigaci\u00f3n primaria \u2014encuestas a consumidores\u2014 para resolverse con mayor certeza.'),
      ]),

      h2('2.4 La ventaja intangible: Marco de Equidad de Marca de Aaker'),

      bp([
        t('Para responder a la pregunta de durabilidad, el Marco de Equidad de Marca de Aaker (1991) permite evaluar los activos intangibles que sostienen la posici\u00f3n competitiva m\u00e1s all\u00e1 del producto. Aplicado a Fenty Beauty, la '),
        CLAVE(),
        t(' en cada dimensi\u00f3n genera los siguientes resultados:'),
      ]),

      new Paragraph({
        children: [new TextRun({ text: 'Tabla 1. Evaluaci\u00f3n de equidad de marca de Fenty Beauty mediante el Marco de Aaker (1991)', font: 'Arial', size: 22, italics: true, bold: true })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 60 },
      }),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2200, 6826],
        rows: [
          aHeader('Dimensi\u00f3n de Aaker', 'An\u00e1lisis aplicado a Fenty Beauty'),
          aRow('Calidad percibida',
            'Alta: f\u00f3rmula del producto y Rihanna como figura de autoridad en el segmento de belleza respaldan la percepci\u00f3n de calidad (DA 1).',
            0),
          aRow('Notoriedad de marca',
            'Alta: segunda marca de celebridad m\u00e1s buscada en Google en 2023, con 21\u202f000 b\u00fasquedas mensuales de su base (DA 2).',
            1),
          aRow('Asociaciones de marca',
            'Fuertes y diferenciadas: inclusividad aut\u00e9ntica, diversidad e identidad de Rihanna. Construidas mediante creatividad pionera dif\u00edcil de replicar retrospectivamente (DA 3, DA 4).',
            0),
          aRow('Lealtad de marca',
            'S\u00f3lida en el n\u00facleo (\u201cFenty Family\u201d), aunque presionada por la creciente competencia de Rare Beauty y Rhode (DA 1).',
            1),
          aRow('Implicaci\u00f3n estrat\u00e9gica',
            'La ventaja no reside en el producto \u2014imitable\u2014 sino en activos intangibles construidos v\u00eda creatividad pionera. Las asociaciones y la autenticidad son m\u00e1s dif\u00edciles de copiar que una gama de tonos (DA 1, DA 4).',
            0),
        ],
      }),

      bp([
        t('La implicaci\u00f3n central del an\u00e1lisis de Aaker es que la '),
        CLAVE(),
        t(' de Fenty Beauty oper\u00f3 en dos horizontes: en el corto plazo, gener\u00f3 ventajas funcionales visibles; en el largo plazo, construy\u00f3 activos intangibles de marca \u2014asociaciones, notoriedad, autenticidad\u2014 que son sustancialmente m\u00e1s dif\u00edciles de imitar. Esta distincci\u00f3n permite reconciliar la evidencia de imitaci\u00f3n funcional con la continuidad del liderazgo financiero de la marca '),
        tb('(DA 2, DA 4)', { size: 24 }),
        t('.'),
      ], 120, 80),

      new Paragraph({ children: [new PageBreak()], spacing: { before: 0, after: 0 } }),

      // ════════════════════════════════════════════════════════════════════
      // 3. CONCLUSI\u00d3N
      // ════════════════════════════════════════════════════════════════════
      h1('3. Conclusi\u00f3n'),

      bp([
        t('La pregunta de investigaci\u00f3n planteada era: '),
        ti('\u00bfde qu\u00e9 manera la '),
        new TextRun({ text: 'creatividad', font: 'Arial', size: 24, italics: true, bold: true }),
        ti(' en la estrategia de marketing inclusivo impact\u00f3 en el posicionamiento de Fenty Beauty en la industria cosm\u00e9tica?'),
        t(' La evidencia analizada permite responder que la '),
        CLAVE(),
        t(' impact\u00f3 de manera positiva, diferenciada y estructuralmente profunda el posicionamiento de la marca, operando en dos horizontes temporales distintos.'),
      ]),

      bp([
        t('En el corto plazo, la '),
        CLAVE(),
        t(' expresada en las 4Ps \u2014especialmente en la distribuci\u00f3n equitativa de tonos, las decisiones de promoci\u00f3n no convencionales y el lanzamiento global simult\u00e1neo\u2014 permiti\u00f3 a Fenty Beauty identificar y ocupar creativamente un cuadrante vac\u00edo del mapa de posicionamiento y construir una ventaja de diferenciaci\u00f3n de Porter medible. La evidencia cuantitativa respalda este impacto: $602,4 millones en ingresos en 2023 y liderazgo entre marcas de belleza de celebridad '),
        tb('(DA 2)', { size: 24 }),
        t('. El \u201cFenty Effect\u201d, que oblig\u00f3 a toda la industria a reformular sus est\u00e1ndares, es la prueba cualitativa m\u00e1s contundente de que la '),
        CLAVE(),
        t(' de Fenty Beauty transform\u00f3 el mercado en su conjunto.'),
      ]),

      bp([
        t('Sin embargo, el an\u00e1lisis de Porter revela que las ventajas funcionales derivadas de la '),
        CLAVE(),
        t(' son imitables: la industria respondi\u00f3 con rapidez y para 2024 la inclusividad de gama ya no es exclusiva de Fenty Beauty '),
        tb('(DA 1)', { size: 24 }),
        t('. En el largo plazo, la ventaja m\u00e1s duradera reside en la equidad de marca \u2014evaluada mediante el Marco de Aaker\u2014: las asociaciones de autenticidad e identidad cultural, construidas a trav\u00e9s de una '),
        CLAVE(),
        t(' pionera que apel\u00f3 a segmentos ignorados '),
        tb('(DA 3, DA 4)', { size: 24 }),
        t(', son activos intangibles dif\u00edcilmente replicables por competidores que adoptaron la inclusividad de forma reactiva.'),
      ]),

      bp([
        t('Esta investigaci\u00f3n presenta limitaciones relevantes: se basa exclusivamente en fuentes secundarias y no captura percepciones directas de consumidores. Permanece sin resolver c\u00f3mo Fenty Beauty puede sostener su '),
        CLAVE(),
        t(' diferenciadora ante competidores emergentes en un mercado donde la inclusividad se ha convertido en est\u00e1ndar; una pregunta que requerir\u00eda investigaci\u00f3n primaria y datos m\u00e1s recientes para abordarse con mayor rigor.'),
      ]),

      new Paragraph({ children: [new PageBreak()], spacing: { before: 0, after: 0 } }),

      // ════════════════════════════════════════════════════════════════════
      // BIBLIOGRAF\u00cdA — DA primero (req. G checklist)
      // ════════════════════════════════════════════════════════════════════
      h1('Bibliograf\u00eda'),

      new Paragraph({
        children: [tb('Documentos de apoyo', { size: 22, color: '1A3C5E' })],
        spacing: { before: 80, after: 60 },
      }),

      ...[
        'DA 1 \u2014 Kravchuk, I. & Krupskyi, O. P. (2025). Inclusive branding as the basis for Fenty Beauty\u2019s strategic positioning in the global beauty industry. ResearchGate / MDPI. https://www.researchgate.net/publication/393262706',
        'DA 2 \u2014 Cosmetify & Beauty Packaging / Upbeat Agency (2023). Leading celebrity beauty brands worldwide in 2023, by revenue. Statista. https://www.statista.com/statistics/1372858/leading-celebrity-beauty-brands-worldwide-by-revenue/',
        'DA 3 \u2014 Fenty Beauty. (2024). About Fenty Beauty: Beauty for All [Sitio web oficial \u2014 fuente interna]. LVMH. https://www.fentybeauty.com/pages/about-us',
        'DA 4 \u2014 Brown, P., Burns, T., Harris, T., Lucas, C. & Zizaoui, I. (2022, 8 de febrero). The rise of the inclusive consumer. McKinsey & Company. https://www.mckinsey.com/industries/retail/our-insights/the-rise-of-the-inclusive-consumer',
        'DA 5 \u2014 Hudson, S., Pacchia, M., Weaver, K. & Berg, A. (2023). The beauty market in 2023: A special State of Fashion report. McKinsey & Company. https://www.mckinsey.com/industries/retail/our-insights/the-beauty-market-in-2023-a-special-state-of-fashion-report',
      ].map(ref => new Paragraph({
        children: [t(ref, { size: 22 })],
        spacing: { before: 60, after: 60 },
        alignment: AlignmentType.JUSTIFIED,
        indent: { left: 720, hanging: 720 },
      })),

      new Paragraph({
        children: [tb('Fuentes adicionales (herramientas y teor\u00edas)', { size: 22, color: '1A3C5E' })],
        spacing: { before: 160, after: 60 },
      }),

      ...[
        'Aaker, D. A. (1991). Managing brand equity: Capitalizing on the value of a brand name. Free Press.',
        'IB. (2022). Gu\u00eda de Gesti\u00f3n Empresarial. International Baccalaureate Organization.',
        'Kotler, P. & Armstrong, G. (2021). Principles of marketing (18th ed.). Pearson.',
        'Porter, M. E. (1985). Competitive advantage: Creating and sustaining superior performance. Free Press.',
      ].map(ref => new Paragraph({
        children: [t(ref, { size: 22 })],
        spacing: { before: 60, after: 60 },
        alignment: AlignmentType.JUSTIFIED,
        indent: { left: 720, hanging: 720 },
      })),

      new Paragraph({ children: [new PageBreak()], spacing: { before: 0, after: 0 } }),

      // ════════════════════════════════════════════════════════════════════
      // GU\u00cdA DE DOCUMENTOS DE APOYO
      // ════════════════════════════════════════════════════════════════════
      h1('Gu\u00eda de Documentos de Apoyo'),

      bp([t('Esta p\u00e1gina es una gu\u00eda operativa para imprimir, destacar y compilar los cinco documentos de apoyo en un \u00fanico PDF para su carga al IB. No cuenta para el c\u00f3mputo de palabras.')]),

      ...[
        {
          label: 'DA 1 \u2014 Art\u00edculo acad\u00e9mico (fuente externa)',
          title: 'Kravchuk, I. & Krupskyi, O. P. (2025). "Inclusive Branding as The Basis for Fenty Beauty\'s Strategic Positioning in the Global Beauty Industry." ResearchGate.',
          url: 'https://www.researchgate.net/publication/393262706',
          type: 'Art\u00edculo acad\u00e9mico \u2014 investigaci\u00f3n cualitativa con revisi\u00f3n de pares',
          highlight: 'Destacar: secci\u00f3n de diferenciaci\u00f3n estrat\u00e9gica, conclusi\u00f3n sobre autenticidad como ventaja y el dato del 77\u00a0% de tonos concentrados en pieles claras por la competencia. Tambi\u00e9n la menci\u00f3n al "Fenty Effect" y a Rare Beauty/Rhode como competidores (2024).',
        },
        {
          label: 'DA 2 \u2014 Datos cuantitativos (fuente externa)',
          title: 'Cosmetify & Beauty Packaging / Upbeat Agency (2023). "Leading Celebrity Beauty Brands Worldwide in 2023, by Revenue." Statista.',
          url: 'https://www.statista.com/statistics/1372858/leading-celebrity-beauty-brands-worldwide-by-revenue/',
          type: 'Base de datos cuantitativa \u2014 estad\u00edstica con fuente Statista',
          highlight: 'Destacar: el gr\u00e1fico de barras con el ranking completo y el dato de $602,4 millones USD de Fenty Beauty vs. Anomaly ($542,7M) y Kylie Cosmetics ($380,4M). Tambi\u00e9n: \u201ccelebrity beauty brands generated over $1 billion in 2023\u201d.',
        },
        {
          label: 'DA 3 \u2014 Sitio web oficial de Fenty Beauty (fuente INTERNA)',
          title: 'Fenty Beauty. (2024). About Fenty Beauty: Beauty for All. LVMH / Kendo Holdings.',
          url: 'https://www.fentybeauty.com/pages/about-us',
          type: 'Fuente interna \u2014 comunicaci\u00f3n oficial de la organizaci\u00f3n objeto de estudio',
          highlight: 'Destacar: la declaraci\u00f3n de misi\u00f3n (\u201cBeauty for All\u201d), la menci\u00f3n a la experiencia personal de Rihanna como origen de la idea, y el rango de 50 tonos de base. Esta es la \u00fanica fuente interna y provee la voz directa de la empresa.',
        },
        {
          label: 'DA 4 \u2014 Informe de consultora (fuente externa)',
          title: 'Brown, P., Burns, T., Harris, T., Lucas, C. & Zizaoui, I. (2022, 8 de febrero). "The Rise of the Inclusive Consumer." McKinsey & Company.',
          url: 'https://www.mckinsey.com/industries/retail/our-insights/the-rise-of-the-inclusive-consumer',
          type: 'Informe de consultora global \u2014 investigaci\u00f3n cuantitativa (encuesta n=270 millones)',
          highlight: 'Destacar: dato principal \u201ctwo out of three Americans told us their social values now shape their shopping choices\u201d y el Exhibit 1 sobre la capacidad del consumidor inclusivo de influir en todos los grupos demogr\u00e1ficos. Tambi\u00e9n: preferencia por recomendaciones de pares vs. publicidad.',
        },
        {
          label: 'DA 5 \u2014 Informe de mercado (fuente externa)',
          title: 'Hudson, S., Pacchia, M., Weaver, K. & Berg, A. (2023). "The Beauty Market in 2023: A Special State of Fashion Report." McKinsey & Company.',
          url: 'https://www.mckinsey.com/industries/retail/our-insights/the-beauty-market-in-2023-a-special-state-of-fashion-report',
          type: 'Informe de mercado sectorial \u2014 McKinsey & The Business of Fashion',
          highlight: 'Destacar: Exhibit 1 con proyecci\u00f3n del mercado cosm\u00e9tico ($430B en 2022 \u2192 $580B para 2027), secci\u00f3n sobre saturaci\u00f3n del mercado y el dato de que solo 5 de 46 marcas independientes lanzadas post-2005 superaron $250M de ventas para 2022.',
        },
      ].flatMap(da => [
        new Paragraph({
          children: [new TextRun({ text: da.label, font: 'Arial', size: 24, bold: true, color: '1A3C5E' })],
          spacing: { before: 240, after: 60 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 3, color: '1A3C5E', space: 1 } },
        }),
        new Table({
          width: { size: 9026, type: WidthType.DXA },
          columnWidths: [2000, 7026],
          rows: [
            ['Fuente completa', da.title],
            ['URL', da.url],
            ['Tipo de fuente', da.type],
            ['Fragmentos a destacar', da.highlight],
          ].map(([k, v], i) =>
            new TableRow({
              children: [
                new TableCell({
                  borders: allB,
                  width: { size: 2000, type: WidthType.DXA },
                  shading: { fill: i % 2 === 0 ? 'F0F4FA' : 'FFFFFF', type: ShadingType.CLEAR },
                  margins: { top: 60, bottom: 60, left: 100, right: 100 },
                  children: [new Paragraph({ children: [tb(k, { size: 20 })], spacing: { before: 0, after: 0 } })],
                }),
                new TableCell({
                  borders: allB,
                  width: { size: 7026, type: WidthType.DXA },
                  shading: { fill: i % 2 === 0 ? 'F0F4FA' : 'FFFFFF', type: ShadingType.CLEAR },
                  margins: { top: 60, bottom: 60, left: 100, right: 100 },
                  children: [new Paragraph({
                    children: [t(v, { size: 20 })],
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 0 },
                  })],
                }),
              ],
            })
          ),
        }),
      ]),

    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/mnt/user-data/outputs/Fenty_Beauty_IA_v2_Luciana_Carlo.docx', buf);
  console.log('Done.');
});
