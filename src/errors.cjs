'use strict';

// Mensagens publicas contem somente codigos, nunca texto do documento.
class Blocked extends Error {
  constructor(code) { super(code); this.name = 'Blocked'; this.code = code; }
}
function requireCondition(condition, code) { if (!condition) throw new Blocked(code); }
module.exports = { Blocked, requireCondition };
