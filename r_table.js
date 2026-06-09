/**
 * Database r_tabel Pearson untuk Signifikansi 5% dan 1% (Uji Dua Arah / Two-Tailed)
 * Berdasarkan Degrees of Freedom (df) = N - 2
 */
const rTableData = {
  // df: [r_0.05, r_0.01]
  1: [0.9969, 0.9999],
  2: [0.9500, 0.9900],
  3: [0.8783, 0.9587],
  4: [0.8114, 0.9172],
  5: [0.7545, 0.8745],
  6: [0.7067, 0.8343],
  7: [0.6664, 0.7977],
  8: [0.6319, 0.7646],
  9: [0.6021, 0.7348],
  10: [0.5760, 0.7079],
  11: [0.5529, 0.6835],
  12: [0.5324, 0.6614],
  13: [0.5140, 0.6411],
  14: [0.4973, 0.6226],
  15: [0.4821, 0.6055],
  16: [0.4683, 0.5897],
  17: [0.4555, 0.5751],
  18: [0.4438, 0.5614],
  19: [0.4329, 0.5487],
  20: [0.4227, 0.5368],
  21: [0.4132, 0.5256],
  22: [0.4044, 0.5151],
  23: [0.3961, 0.5052],
  24: [0.3882, 0.4958],
  25: [0.3809, 0.4869],
  26: [0.3739, 0.4785],
  27: [0.3673, 0.4705],
  28: [0.3610, 0.4629],
  29: [0.3550, 0.4556],
  30: [0.3494, 0.4487],
  31: [0.3440, 0.4421],
  32: [0.3388, 0.4357],
  33: [0.3338, 0.4296],
  34: [0.3291, 0.4238],
  35: [0.3246, 0.4182],
  36: [0.3202, 0.4128],
  37: [0.3160, 0.4076],
  38: [0.3120, 0.4026],
  39: [0.3081, 0.3978],
  40: [0.3044, 0.3932],
  41: [0.3008, 0.3887],
  42: [0.2973, 0.3843],
  43: [0.2940, 0.3801],
  44: [0.2907, 0.3761],
  45: [0.2876, 0.3721],
  46: [0.2845, 0.3683],
  47: [0.2816, 0.3646],
  48: [0.2787, 0.3610],
  49: [0.2759, 0.3575],
  50: [0.2732, 0.3542],
  60: [0.2500, 0.3248],
  70: [0.2319, 0.3017],
  80: [0.2172, 0.2830],
  90: [0.2050, 0.2673],
  100: [0.1946, 0.2540],
  125: [0.1743, 0.2278],
  150: [0.1593, 0.2083],
  200: [0.1381, 0.1808],
  300: [0.1129, 0.1479],
  400: [0.0978, 0.1282],
  500: [0.0875, 0.1147],
  1000: [0.0619, 0.0812]
};

window.rTable = {
  /**
   * Mengambil nilai r_tabel berdasarkan df dan level signifikansi (0.05 atau 0.01)
   * Jika df tidak ada di tabel, dilakukan interpolasi linier atau pendekatan rumus.
   * @param {number} df Degrees of Freedom (N - 2)
   * @param {number} alpha Signifikansi (0.05 atau 0.01)
   * @returns {number} Nilai r kritis
   */
  get: function(df, alpha = 0.05) {
    if (df < 1) return 0;
    const idx = (alpha === 0.01) ? 1 : 0;

    // Jika df ada langsung di database
    if (rTableData[df]) {
      return rTableData[df][idx];
    }

    // Jika df melebihi batas terbesar (1000), gunakan rumus pendekatan
    if (df > 1000) {
      const z = (alpha === 0.01) ? 2.5758 : 1.95996;
      return z / Math.sqrt(df + Math.pow(z, 2));
    }

    // Jika di antara entri, cari intervalnya dan lakukan interpolasi linier
    const dfs = Object.keys(rTableData).map(Number).sort((a, b) => a - b);
    let lowerDf = dfs[0];
    let upperDf = dfs[dfs.length - 1];

    for (let i = 0; i < dfs.length; i++) {
      if (dfs[i] > df) {
        upperDf = dfs[i];
        lowerDf = dfs[i - 1];
        break;
      }
    }

    const rLower = rTableData[lowerDf][idx];
    const rUpper = rTableData[upperDf][idx];

    // Rumus interpolasi linier
    const fraction = (df - lowerDf) / (upperDf - lowerDf);
    const rVal = rLower + fraction * (rUpper - rLower);

    return parseFloat(rVal.toFixed(4));
  }
};
