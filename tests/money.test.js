import test from 'node:test';import assert from 'node:assert/strict';
import {roundMoney,moneyCents} from '../src/money.js';
test('rounds extra decimals to nearest piastre',()=>{assert.equal(roundMoney('934.6199'),934.62);assert.equal(roundMoney('3000.145'),3000.15);assert.equal(roundMoney('250.124'),250.12);});
test('carries 99 mills into the next piastre/pound correctly',()=>{assert.equal(roundMoney('150.999'),151);assert.equal(moneyCents('934.6199'),93462);});
