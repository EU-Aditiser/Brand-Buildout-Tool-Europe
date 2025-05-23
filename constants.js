// Client ID and API key from the Developer Console
var CLIENT_ID = '584505830851-0k3n478q7ecobi5k97o460c4dqoj8jf5.apps.googleusercontent.com';
var API_KEY = 'AIzaSyC6LwQLz7R0OfeaXQyG5lIglq_F7Xo_pI8';

// Array of API discovery doc URLs for APIs used by the quickstart
var DISCOVERY_DOCS = ["https://sheets.googleapis.com/$discovery/rest?version=v4"];

// Authorization scopes required by the API; multiple scopes can be
// included, separated by spaces.
var SCOPES = "https://www.googleapis.com/auth/spreadsheets";

const PRODUCT_KEYWORDS = [
  {
      prefix: "",
      suffix: ""
    },
    {
      prefix: "",
      suffix: " products"
    },
    {
      prefix: "buy ",
      suffix: ""
    },
    {
      prefix: "buy ",
      suffix: " online"
    },
    {
      prefix: "cheap ",
      suffix: ""
    },
    {
      prefix: "",
      suffix: " coupons"
    },
    {
      prefix: "",
      suffix: " discount"
    },
    {
      prefix: "",
      suffix: " online"
    },
    {
      prefix: "",
      suffix: " online store"
    },
    {
      prefix: "",
      suffix: " sale"
    }
  ];

  const SIMPLE_DOMAIN_KEYWORDS = [
  {
    prefix: "www ",
    suffix: " com"
  },
  {
    prefix: "www",
    suffix: "com"
  },
  {
    prefix: "www ",
    suffix: ""
  },
  {
    prefix: "www",
    suffix: ""
  },
  {
    prefix: "",
    suffix: " com"
  },
  {
    prefix: "",
    suffix: "com"
  }
]

const COMPOUND_DOMAIN_KEYWORDS = [
  {
    prefix: "www ",
    suffix: " com"
  },
  {
    prefix: "www ",
    suffix: ""
  },
  {
    prefix: "",
    suffix: " com"
  },
]

const NEGATIVE_KEYWORDS = [
  {
    keyword: "ingredient",
    match_type: "Negative Broad"
  },
  {
    keyword: "ingredients",
    match_type: "Negative Broad"
  },
  {
    keyword: "address",
    match_type: "Negative Broad"
  },
  {
    keyword: "location",
    match_type: "Negative Broad"
  },
  {
    keyword: "locations",
    match_type: "Negative Broad"
  },
  {
    keyword: "logo",
    match_type: "Negative Broad"
  },
  {
    keyword: "side effect",
    match_type: "Negative Phrase"
  },
  {
    keyword: "side effects",
    match_type: "Negative Phrase"
  },
  {
    keyword: "review",
    match_type: "Negative Broad"
  },
  {
    keyword: "reviews",
    match_type: "Negative Broad"
  }
];