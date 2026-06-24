export type Box = { x: number; y: number; w: number; h: number };

type LayoutArchetype = {
  templateName: string;
  title?: Box;
  body?: Box;
  object?: Box;
  source?: Box;
  number?: Box;
  numberLabel?: Box;
  objects?: readonly Box[];
  notes: string;
};

export const layouts = {
  page: { width: 13.333, height: 7.5 },
  archetypes: {
    titleSlide: {
      templateName: "Title Slide",
      title: { x: 1.411, y: 2.648, w: 10.635, h: 2.696 },
      body: { x: 1.412, y: 5.959, w: 10.635, h: 0.629 },
      notes: "Opening slide: wordmark top-left, large title left, short subtitle lower-left, pale blue field right."
    },
    titleOnly: {
      templateName: "Title only",
      title: { x: 1.351, y: 0.915, w: 10.637, h: 1.025 },
      notes: "Content slide with one dominant visual added by renderer."
    },
    textObject: {
      templateName: "Text + Object",
      title: { x: 1.351, y: 0.915, w: 10.637, h: 0.705 },
      body: { x: 1.391, y: 1.94, w: 4.95, h: 4.647 },
      object: { x: 6.779, y: 1.94, w: 5.244, h: 4.647 },
      notes: "Short text claim with one visual/object."
    },
    text: {
      templateName: "Text",
      title: { x: 1.351, y: 0.915, w: 10.637, h: 0.705 },
      body: { x: 1.391, y: 1.94, w: 10.637, h: 4.647 },
      notes: "Narrative slide; avoid for metrics unless needed."
    },
    accentTitle: {
      templateName: "Accent title",
      title: { x: 1.351, y: 0.915, w: 10.637, h: 0.705 },
      notes: "Blue title treatment."
    },
    textNumber: {
      templateName: "Text + Number",
      title: { x: 1.351, y: 0.915, w: 10.637, h: 0.705 },
      body: { x: 1.391, y: 1.94, w: 6.482, h: 4.647 },
      number: { x: 8.487, y: 2.688, w: 3.537, h: 1.862 },
      numberLabel: { x: 8.485, y: 4.623, w: 3.539, h: 0.503 },
      notes: "One short claim plus one large KPI."
    },
    oneObject: {
      templateName: "One Object",
      title: { x: 1.351, y: 0.915, w: 10.637, h: 0.705 },
      object: { x: 1.351, y: 1.94, w: 10.637, h: 5.562 },
      notes: "Single dominant chart/object."
    },
    objectTitle: {
      templateName: "Object + Title",
      title: { x: 1.351, y: 0.915, w: 10.637, h: 0.705 },
      object: { x: 1.351, y: 1.94, w: 10.637, h: 5.562 },
      notes: "Standard analytical chart slide."
    },
    twoTexts: {
      templateName: "2 Texts",
      title: { x: 1.351, y: 0.915, w: 10.637, h: 0.705 },
      objects: [
        { x: 1.391, y: 1.94, w: 5.166, h: 4.647 },
        { x: 6.779, y: 1.94, w: 5.244, h: 4.647 }
      ],
      notes: "Two balanced text blocks."
    },
    twoNumbers: {
      templateName: "2 numbers",
      title: { x: 1.351, y: 0.915, w: 10.637, h: 0.705 },
      notes: "Two KPI summary layout."
    },
    accentIdea: {
      templateName: "Accent Idea",
      notes: "Large accent statement."
    },
    oneFourNumbers: {
      templateName: "1_4 numbers",
      notes: "One lead KPI plus four supporting numbers."
    },
    finalSlideQr2: {
      templateName: "Final slide + QR2",
      notes: "Closing slide with QR."
    },
    divider: {
      templateName: "Divider",
      title: { x: 1.391, y: 3.051, w: 5.914, h: 2.272 },
      body: { x: 1.391, y: 1.939, w: 3.389, h: 0.705 },
      notes: "Full blue section break with white title."
    },
    emptySlide: {
      templateName: "Empty slide",
      notes: "Blank template-controlled slide."
    },
    accentTitleBright: {
      templateName: "Accent title bright",
      notes: "Bright accent title variation."
    },
    twoObjects: {
      templateName: "2 Objects",
      title: { x: 1.351, y: 0.915, w: 10.637, h: 0.705 },
      objects: [
        { x: 1.391, y: 1.94, w: 5.166, h: 5.562 },
        { x: 6.779, y: 1.94, w: 5.244, h: 5.562 }
      ],
      notes: "Two side-by-side analytical visuals with comparable weight."
    },
    fourObjects: {
      templateName: "4 Objects",
      title: { x: 1.351, y: 0.915, w: 10.637, h: 0.705 },
      notes: "Four-object grid."
    },
    oneIdea: {
      templateName: "One Idea",
      notes: "One statement or idea only."
    },
    quote: {
      templateName: "Quote",
      notes: "Quote slide."
    },
    accentQuote: {
      templateName: "Accent Quote",
      notes: "Accent quote slide."
    },
    threeNumbers: {
      templateName: "3 numbers",
      title: { x: 1.351, y: 0.915, w: 10.637, h: 0.705 },
      notes: "Three KPI summary layout."
    },
    fullMockup: {
      templateName: "Full Mockup",
      notes: "Full-screen product mockup."
    },
    topMockup: {
      templateName: "Top Mockup",
      notes: "Top-aligned product mockup."
    },
    bottomMockup: {
      templateName: "Bottom Mockup",
      notes: "Bottom-aligned product mockup."
    },
    finalSlide: {
      templateName: "Final slide",
      notes: "Simple closing slide."
    }
  } satisfies Record<string, LayoutArchetype>,
  derived: {
    content: {
      title: { x: 1.351, y: 0.915, w: 10.637, h: 0.705 },
      body: { x: 1.391, y: 1.94, w: 4.95, h: 4.647 },
      image: { x: 6.779, y: 1.94, w: 5.244, h: 4.647 }
    },
    comparison: {
      title: { x: 1.351, y: 0.915, w: 10.637, h: 0.705 },
      body: { x: 1.391, y: 1.62, w: 10.637, h: 0.28 },
      leftImage: { x: 1.391, y: 1.94, w: 5.166, h: 5.0 },
      rightImage: { x: 6.779, y: 1.94, w: 5.244, h: 5.0 }
    },
    chartSlide: {
      title: { x: 1.351, y: 0.915, w: 7.0, h: 0.705 },
      body: { x: 1.391, y: 1.56, w: 10.637, h: 0.3 },
      kpi1: { x: 8.8, y: 0.915, w: 2.0, h: 0.6 },
      kpi2: { x: 10.8, y: 0.915, w: 2.0, h: 0.6 },
      chart: { x: 0.98, y: 2.05, w: 11.25, h: 4.95 },
      source: { x: 1.351, y: 7.05, w: 9.2, h: 0.14 }
    }
  }
} as const;

export type Layouts = typeof layouts;
