import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { parseCyberbizOrderEmail } from '../parser.js';

describe('parseCyberbizOrderEmail', () => {
  it('should parse a basic Cyberbiz order email', () => {
    const email = {
      id: 'msg-001',
      subject: '您有一筆新訂單 訂單編號: C20260101001',
      date: 'Wed, 01 Jan 2026 10:00:00 +0800',
      body: `
        <html><body>
          <p>訂單編號：C20260101001</p>
          <p>訂單日期：2026/01/01 10:00</p>
          <p>收件人：王小明</p>
          <p>電話：0912-345-678</p>
          <p>地址：台北市信義區信義路五段7號</p>
          <p>付款方式：信用卡</p>
          <p>配送方式：宅配</p>
          <table>
            <tr><th>商品</th><th>數量</th><th>金額</th></tr>
            <tr><td>能量膠 - 柑橘口味</td><td>3</td><td>NT$450</td></tr>
            <tr><td>電解質粉 - 檸檬口味</td><td>1</td><td>NT$350</td></tr>
          </table>
          <p>小計：NT$800</p>
          <p>運費：NT$60</p>
          <p>折扣：-NT$100</p>
          <p>總計：NT$760</p>
        </body></html>
      `,
    };

    const order = parseCyberbizOrderEmail(email);

    assert.notEqual(order, null);
    assert.equal(order.orderNumber, 'C20260101001');
    assert.equal(order.orderDate, '2026/01/01 10:00');
    assert.equal(order.customerName, '王小明');
    assert.equal(order.customerPhone, '0912-345-678');
    assert.equal(order.shippingAddress, '台北市信義區信義路五段7號');
    assert.equal(order.paymentMethod, '信用卡');
    assert.equal(order.shippingMethod, '宅配');
    assert.equal(order.items.length, 2);
    assert.equal(order.items[0].name, '能量膠 - 柑橘口味');
    assert.equal(order.items[0].quantity, 3);
    assert.equal(order.items[0].price, 450);
    assert.equal(order.subtotal, 800);
    assert.equal(order.shippingFee, 60);
    assert.equal(order.discount, 100);
    assert.equal(order.totalAmount, 760);
  });

  it('should extract order number from subject with # prefix', () => {
    const email = {
      id: 'msg-002',
      subject: '新訂單通知 #C20260201002',
      date: 'Sun, 01 Feb 2026 12:00:00 +0800',
      body: '<html><body><p>總計：NT$1,200</p></body></html>',
    };

    const order = parseCyberbizOrderEmail(email);
    assert.notEqual(order, null);
    assert.equal(order.orderNumber, 'C20260201002');
    assert.equal(order.totalAmount, 1200);
  });

  it('should detect order status from subject', () => {
    const email = {
      id: 'msg-003',
      subject: '訂單編號: C123456 已出貨',
      date: 'Mon, 02 Feb 2026 09:00:00 +0800',
      body: '<html><body><p>出貨通知</p></body></html>',
    };

    const order = parseCyberbizOrderEmail(email);
    assert.notEqual(order, null);
    assert.equal(order.orderStatus, '已出貨');
  });

  it('should return null for email without order number', () => {
    const email = {
      id: 'msg-004',
      subject: '一般通知信件',
      date: 'Mon, 02 Feb 2026 09:00:00 +0800',
      body: '<html><body><p>Hello</p></body></html>',
    };

    const order = parseCyberbizOrderEmail(email);
    assert.equal(order, null);
  });

  it('should parse items from text pattern (name x quantity $price)', () => {
    const email = {
      id: 'msg-005',
      subject: '訂單編號: C999888',
      date: 'Tue, 03 Feb 2026 15:00:00 +0800',
      body: `
        <html><body>
          <p>能量棒 x2 $300</p>
          <p>運動飲料 x3 $450</p>
          <p>總計：$750</p>
        </body></html>
      `,
    };

    const order = parseCyberbizOrderEmail(email);
    assert.notEqual(order, null);
    assert.equal(order.items.length, 2);
    assert.equal(order.items[0].name, '能量棒');
    assert.equal(order.items[0].quantity, 2);
    assert.equal(order.items[1].quantity, 3);
  });
});
