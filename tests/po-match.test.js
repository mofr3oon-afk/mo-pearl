import test from 'node:test';import assert from 'node:assert/strict';
import {extractPo,comparePo} from '../src/po-match.js';
test('reads several PO labels',()=>{for(const txt of ['PO: 4300228298','Purchase Order Number: 4300228298','P.O. 4300228298'])assert.equal(extractPo(txt).number,'4300228298');});
test('matching PO passes',()=>assert.equal(comparePo('PO: 4300228298 EXPRO EGYPT LLC',{name:'EXPRO EGYPT LLC',po:'4300228298'},{}).checks[0].kind,'ok'));
test('mismatching PO is an error',()=>assert.equal(comparePo('PO: 4300228299',{po:'4300228298'},{}).status,'error'));
test('different PO total is review not a hard error',()=>assert.notEqual(comparePo('PO: 4300228298 ORDER TOTAL: 2000',{po:'4300228298'},{total:500}).status,'error'));
