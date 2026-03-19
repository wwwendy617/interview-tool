// API helper
const api = {
  async getAll() {
    const res = await fetch('/api/interviews');
    return res.json();
  },

  async getOne(id) {
    const res = await fetch(`/api/interviews/${id}`);
    return res.json();
  },

  async create(data) {
    const res = await fetch('/api/interviews', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    });
    return res.json();
  },

  async update(id, data) {
    const res = await fetch(`/api/interviews/${id}`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    });
    return res.json();
  },

  async delete(id) {
    const res = await fetch(`/api/interviews/${id}`, {
      method: 'DELETE'
    });
    return res.json();
  }
};
