const backgroundColor = {
  red: 239 / 255,
  green: 239 / 255,
  blue: 239 / 255,
};

const borderStyle = {
  style: "SOLID",
  color: {
    red: 0.0,
    green: 0.0,
    blue: 0.0,
    alpha: 1.0,
  },
};

const firstColumn = {
  cell: {
    userEnteredFormat: {
      backgroundColor,
    },
  },
  fields: "userEnteredFormat(backgroundColor)",
};

export const styles = {
  borders: {
    top: borderStyle,
    right: borderStyle,
    bottom: borderStyle,
    left: borderStyle,
  },
  header: {
    cell: {
      userEnteredFormat: {
        backgroundColor,
        horizontalAlignment: "CENTER",
        textFormat: {
          bold: true,
        },
      },
    },
    fields: "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
  },
  firstColumn,
  footer: firstColumn,
};
