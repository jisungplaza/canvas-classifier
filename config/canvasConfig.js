module.exports = {
  sizeMap: [
    { w: 18.0, h: 14.0, code: "0F" },

    { w: 22.7, h: 15.8, code: "1F" },
    { w: 22.7, h: 14.0, code: "1P" },
    { w: 22.7, h: 12.0, code: "1M" },
    { w: 15.8, h: 15.8, code: "1S" },

    { w: 25.8, h: 17.9, code: "2F" },
    { w: 25.8, h: 16.0, code: "2P" },
    { w: 25.8, h: 14.0, code: "2M" },
    { w: 17.9, h: 17.9, code: "2S" },

    { w: 27.3, h: 22.0, code: "3F" },
    { w: 27.3, h: 19.0, code: "3P" },
    { w: 27.3, h: 16.0, code: "3M" },
    { w: 22.0, h: 22.0, code: "3S" },

    { w: 33.4, h: 24.2, code: "4F" },
    { w: 33.4, h: 21.2, code: "4P" },
    { w: 33.4, h: 19.0, code: "4M" },
    { w: 24.2, h: 24.2, code: "4S" },

    { w: 34.8, h: 27.3, code: "5F" },
    { w: 34.8, h: 24.2, code: "5P" },
    { w: 34.8, h: 21.2, code: "5M" },
    { w: 27.3, h: 27.3, code: "5S" },

    { w: 40.9, h: 31.8, code: "6F" },
    { w: 40.9, h: 27.3, code: "6P" },
    { w: 40.9, h: 24.2, code: "6M" },
    { w: 31.8, h: 31.8, code: "6S" },

    { w: 45.5, h: 37.9, code: "8F" },
    { w: 45.5, h: 33.4, code: "8P" },
    { w: 45.5, h: 27.3, code: "8M" },
    { w: 37.9, h: 37.9, code: "8S" },

    { w: 53.0, h: 45.5, code: "10F" },
    { w: 53.0, h: 40.9, code: "10P" },
    { w: 53.0, h: 33.4, code: "10M" },
    { w: 45.5, h: 45.5, code: "10S" },

    { w: 60.6, h: 50.0, code: "12F" },
    { w: 60.6, h: 45.5, code: "12P" },
    { w: 60.6, h: 40.9, code: "12M" },
    { w: 50.0, h: 50.0, code: "12S" },

    { w: 65.1, h: 53.0, code: "15F" },
    { w: 65.1, h: 50.0, code: "15P" },
    { w: 65.1, h: 45.5, code: "15M" },
    { w: 53.0, h: 53.0, code: "15S" },

    { w: 72.7, h: 60.6, code: "20F" },
    { w: 72.7, h: 53.0, code: "20P" },
    { w: 72.7, h: 50.0, code: "20M" },
    { w: 60.6, h: 60.6, code: "20S" },

    { w: 80.3, h: 65.1, code: "25F" },
    { w: 80.3, h: 60.6, code: "25P" },
    { w: 80.3, h: 53.0, code: "25M" },
    { w: 65.1, h: 65.1, code: "25S" },

    { w: 90.9, h: 72.7, code: "30F" },
    { w: 90.9, h: 65.1, code: "30P" },
    { w: 90.9, h: 60.6, code: "30M" },
    { w: 72.7, h: 72.7, code: "30S" },

    { w: 100.0, h: 80.3, code: "40F" },
    { w: 100.0, h: 72.7, code: "40P" },
    { w: 100.0, h: 65.1, code: "40M" },
    { w: 80.3, h: 80.3, code: "40S" },

    { w: 116.8, h: 91.0, code: "50F" },
    { w: 116.8, h: 80.3, code: "50P" },
    { w: 116.8, h: 72.7, code: "50M" },
    { w: 91.0, h: 91.0, code: "50S" },

    { w: 130.3, h: 97.0, code: "60F" },
    { w: 130.3, h: 89.4, code: "60P" },
    { w: 130.3, h: 80.3, code: "60M" },
    { w: 97.0, h: 97.0, code: "60S" },

    { w: 145.5, h: 112.1, code: "80F" },
    { w: 145.5, h: 97.0, code: "80P" },
    { w: 145.5, h: 89.4, code: "80M" },
    { w: 112.1, h: 112.1, code: "80S" },

    { w: 162.2, h: 130.3, code: "100F" },
    { w: 162.2, h: 112.1, code: "100P" },
    { w: 162.2, h: 97.0, code: "100M" },
    { w: 130.3, h: 130.3, code: "100S" },
  ],

typeRules: [
  // ✅ 판넬 (천 정보 없을 때 기본으로 여기로)
  {
    labelKo: "판넬",
    keywords: [
      "birch panel",
      "birch",
      "paper lable",
      "paper label",
      "total thickness",
      "1.8cm",
      "30*30",
      "30x30",
      // ✅ canvas panel 문구가 들어오되 cotton 정보가 없으면 판넬로 가게끔
      "canvas panel",
      "panel"
    ],
  },

  // 기존 캔버스류
  { labelKo: "프레임 캔버스", keywords: ["frame canvas"] },
  { labelKo: "백스플라인 캔버스", keywords: ["back spline"] },
  { labelKo: "아사 캔버스", keywords: ["linen"] },
  { labelKo: "블랙 캔버스", keywords: ["black"] },

  // ✅ 캔버스보드: typeRules에서는 "천 정보" 키워드만 잡게 해두고,
  // 최종 결정은 index.js에서 panel 계열일 때 cotton 포함이면 캔버스보드로 덮어씀
  {
    labelKo: "캔버스보드",
    keywords: [
      "cotton",
      "gms",
      "gsm",
      "280gms",
      "280 gsm",
      "5309",
      "made in china"
    ],
  },
],


  roundConfig: {
    keywords: ["dia", "ø", "⌀"],
  },

  thicknessMap: [
    { w: 2.9, h: 1.6, labelKo: "정왁구(주황)", nos: [0] },
    { w: 3.5, h: 1.8, labelKo: "정왁구(주황)", nos: [1, 2] },
    { w: 4.0, h: 1.8, labelKo: "정왁구(주황)", nos: [3, 4, 6, 8, 10, 12, 15] },
    { w: 4.5, h: 2.0, labelKo: "정왁구(주황)", nos: [20, 25] },
    { w: 3.8, h: 3.8, labelKo: "정왁구(주황)", nos: [30] },
    { w: 5.7, h: 3.6, labelKo: "정왁구(주황)", nos: [100] },
  ],

  excludedForThickness: [
    "linen",
    "black",
    "back spline",
    "frame canvas",
    "3d",
  ],

  manualOverrides: {},

  sizeTolerance: 0.2,
  thicknessTolerance: 0.15,

  defaultTypeLabel: "일반(파랑)",
};
