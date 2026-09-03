type User = {
  name: string;
  role: 'tester' | 'developer';
};

function greet(user: User): string {
  return `Hello ${user.name}! TypeScript knows your role is ${user.role}.`;
}

const user: User = {
  name: 'Andrew',
  role: 'tester'
};

console.log(greet(user));
console.log('This file was written as TypeScript and compiled to JavaScript before Node.js ran it.');
