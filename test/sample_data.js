const fs = require('fs');
const path = require('path');

function loadSamplePayloads() {
  const dir = path.join(__dirname, '..', 'sample_data');
  return {
    orders: JSON.parse(fs.readFileSync(path.join(dir, 'orders.json'), 'utf8')),
    locations: JSON.parse(fs.readFileSync(path.join(dir, 'locations.json'), 'utf8')),
    products: JSON.parse(fs.readFileSync(path.join(dir, 'products.json'), 'utf8')),
    inventory: JSON.parse(fs.readFileSync(path.join(dir, 'inventory_levels.json'), 'utf8')),
    assignees: JSON.parse(fs.readFileSync(path.join(dir, 'assignees.json'), 'utf8')).sku_assignees
  };
}

module.exports = { loadSamplePayloads };
