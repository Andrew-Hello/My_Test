import os
import platform
import random
import tkinter as tk
from datetime import datetime

CELL=20; WIDTH=500; HEIGHT=500; TICK=110

class Snake:
    def __init__(self, root):
        self.root=root; root.title('Snake - Python')
        self.canvas=tk.Canvas(root,width=WIDTH,height=HEIGHT,bg='#111',highlightthickness=0); self.canvas.pack()
        tk.Label(root,text='Arrow keys / WASD | Enter/Space restart',font=('Segoe UI',11)).pack(pady=6)
        root.bind('<Key>',self.key); root.protocol('WM_DELETE_WINDOW',self.close)
        self.after=None; self.reset()
    def reset(self):
        if self.after: self.root.after_cancel(self.after)
        self.snake=[(8,10),(7,10),(6,10)]; self.direction=(1,0); self.next=(1,0); self.score=0; self.running=True; self.food=self.new_food(); self.draw(); self.after=self.root.after(TICK,self.step)
    def new_food(self):
        while True:
            p=(random.randrange(WIDTH//CELL),random.randrange(HEIGHT//CELL))
            if p not in self.snake: return p
    def key(self,e):
        k=e.keysym.lower()
        if not self.running and k in ('return','space'): self.reset(); return
        m={'up':(0,-1),'w':(0,-1),'down':(0,1),'s':(0,1),'left':(-1,0),'a':(-1,0),'right':(1,0),'d':(1,0)}
        if k in m and (m[k][0]+self.direction[0],m[k][1]+self.direction[1])!=(0,0): self.next=m[k]
    def step(self):
        if not self.running:return
        self.direction=self.next; hx,hy=self.snake[0]; dx,dy=self.direction; head=(hx+dx,hy+dy)
        if head[0]<0 or head[0]>=WIDTH//CELL or head[1]<0 or head[1]>=HEIGHT//CELL or head in self.snake[:-1]: self.running=False; self.draw(); return
        self.snake.insert(0,head)
        if head==self.food: self.score+=1; self.food=self.new_food()
        else:self.snake.pop()
        self.draw(); self.after=self.root.after(TICK,self.step)
    def draw(self):
        c=self.canvas;c.delete('all');fx,fy=self.food;c.create_oval(fx*CELL+3,fy*CELL+3,(fx+1)*CELL-3,(fy+1)*CELL-3,fill='#ff5252',outline='')
        for i,(x,y) in enumerate(self.snake):c.create_rectangle(x*CELL+1,y*CELL+1,(x+1)*CELL-1,(y+1)*CELL-1,fill='#7cff6b' if i==0 else '#38c95d',outline='')
        c.create_text(10,10,anchor='nw',text=f'Score: {self.score}',fill='white',font=('Segoe UI',15,'bold'))
        if not self.running:
            c.create_rectangle(70,190,430,310,fill='#000',outline='white');c.create_text(250,225,text='GAME OVER',fill='white',font=('Segoe UI',24,'bold'));c.create_text(250,270,text=f'Score: {self.score} | Enter/Space restart',fill='white',font=('Segoe UI',11))
    def close(self):
        os.makedirs('output',exist_ok=True)
        with open(os.path.join('output','python_snake_last_run.txt'),'w',encoding='utf-8') as f:
            f.write(f'game=Python Snake\ntime={datetime.now().isoformat()}\nscore={self.score}\npython={platform.python_version()}\nplatform={platform.platform()}\n')
        self.root.destroy()

if __name__=='__main__':
    r=tk.Tk();Snake(r);r.mainloop()
