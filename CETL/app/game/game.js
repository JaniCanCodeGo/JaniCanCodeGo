// Concentration (memory match) game for Clear Enough To Lead.
(function () {
  "use strict";

  var SYMBOLS = ["🌿", "☕", "📚", "🎧", "🧠", "💜", "🌙", "🔥", "🧭", "🎯", "🌊", "⭐"];
  var PAIR_COUNTS = { easy: 6, hard: 12 };
  var BEST_KEY_PREFIX = "cetl-concentration-best-";

  var board = document.getElementById("board");
  var movesEl = document.getElementById("moves");
  var matchesEl = document.getElementById("matches");
  var totalPairsEl = document.getElementById("total-pairs");
  var bestEl = document.getElementById("best");
  var winBanner = document.getElementById("win-banner");
  var winDetail = document.getElementById("win-detail");
  var announcer = document.getElementById("announcer");
  var difficultySelect = document.getElementById("difficulty");
  var newGameBtn = document.getElementById("new-game");

  var state = null;

  function shuffle(items) {
    for (var i = items.length - 1; i > 0; i--) {
      var j = Math.floor(Math.random() * (i + 1));
      var tmp = items[i];
      items[i] = items[j];
      items[j] = tmp;
    }
    return items;
  }

  function readBest(difficulty) {
    try {
      var value = parseInt(localStorage.getItem(BEST_KEY_PREFIX + difficulty), 10);
      return isNaN(value) ? null : value;
    } catch (e) {
      return null;
    }
  }

  function saveBest(difficulty, moves) {
    try {
      localStorage.setItem(BEST_KEY_PREFIX + difficulty, String(moves));
    } catch (e) {
      /* storage unavailable — the game still works, best just isn't kept */
    }
  }

  function announce(message) {
    announcer.textContent = message;
  }

  function newGame() {
    var difficulty = difficultySelect.value;
    var pairCount = PAIR_COUNTS[difficulty] || PAIR_COUNTS.easy;
    var symbols = SYMBOLS.slice(0, pairCount);
    var deck = shuffle(symbols.concat(symbols).slice());

    state = {
      difficulty: difficulty,
      pairCount: pairCount,
      moves: 0,
      matches: 0,
      first: null,
      locked: false
    };

    board.classList.toggle("cols-6", difficulty === "hard");
    board.textContent = "";
    winBanner.hidden = true;
    movesEl.textContent = "0";
    matchesEl.textContent = "0";
    totalPairsEl.textContent = String(pairCount);
    var best = readBest(difficulty);
    bestEl.textContent = best === null ? "—" : String(best) + " moves";

    deck.forEach(function (symbol, index) {
      var card = document.createElement("button");
      card.type = "button";
      card.className = "game-card";
      card.dataset.symbol = symbol;
      card.setAttribute("aria-label", "Card " + (index + 1) + ", face down");
      card.textContent = "?";
      card.addEventListener("click", onCardClick);
      board.appendChild(card);
    });

    announce("New game started with " + pairCount + " pairs.");
  }

  function reveal(card) {
    card.classList.add("flipped");
    card.textContent = card.dataset.symbol;
    card.setAttribute("aria-label", "Card showing " + card.dataset.symbol);
  }

  function hide(card) {
    card.classList.remove("flipped");
    card.textContent = "?";
    card.setAttribute("aria-label", "Card, face down");
  }

  function onCardClick(event) {
    var card = event.currentTarget;
    if (state.locked || card.classList.contains("flipped") || card.classList.contains("matched")) {
      return;
    }

    reveal(card);

    if (!state.first) {
      state.first = card;
      return;
    }

    var first = state.first;
    state.first = null;
    state.moves++;
    movesEl.textContent = String(state.moves);

    if (first.dataset.symbol === card.dataset.symbol) {
      first.classList.add("matched");
      card.classList.add("matched");
      first.disabled = true;
      card.disabled = true;
      state.matches++;
      matchesEl.textContent = String(state.matches);
      announce("Match! " + state.matches + " of " + state.pairCount + " pairs found.");
      if (state.matches === state.pairCount) {
        win();
      }
    } else {
      state.locked = true;
      announce("No match.");
      setTimeout(function () {
        hide(first);
        hide(card);
        state.locked = false;
      }, 900);
    }
  }

  function win() {
    var best = readBest(state.difficulty);
    var isRecord = best === null || state.moves < best;
    if (isRecord) {
      saveBest(state.difficulty, state.moves);
      bestEl.textContent = String(state.moves) + " moves";
    }
    winDetail.textContent =
      state.pairCount + " pairs in " + state.moves + " moves" + (isRecord ? " — a new personal best." : ".");
    winBanner.hidden = false;
    announce("Board cleared in " + state.moves + " moves." + (isRecord ? " New personal best!" : ""));
  }

  newGameBtn.addEventListener("click", newGame);
  difficultySelect.addEventListener("change", newGame);
  newGame();
})();
