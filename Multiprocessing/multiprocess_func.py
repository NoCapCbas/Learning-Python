from typing import Callable, List
"""
Use multiprocessing for CPU-bound tasks that require parallel processing, 
and use multithreading for I/O-bound tasks that involve waiting for 
input/output operations to complete.
"""

def multiprocess_function(func:Callable, args:List) -> None:
	"""
	Multi proccesses, runs each arg on a seperate core
	assuming the pc has 8 cores

	Google to check how many cores your pc has. Most standard pc's 
	have 2-4.
	"""
	import multiprocessing
	pool = multiprocessing.Pool(processes=8)
	pool.map(func, args)

def main(arg):
	"""
	Example Main code here...
	"""
	print(arg)

if __name__ == '__main__':
	# Example list of args passed
	multiprocess_function(main, [1,2,3,4,5])

"""
Multiprocessing is generally preferred over threading in Python for several reasons:

- Performance: Multiprocessing can take advantage of multiple processors or CPU cores on a computer, 
allowing for true parallel processing and potentially significant performance gains. Threading, on 
the other hand, is limited to a single processor or core and can suffer from the Global Interpreter 
Lock (GIL) in Python, which can limit true parallelism.

- Isolation: Each process in multiprocessing has its own memory space, so they are isolated 
from one another. This can prevent issues such as race conditions or deadlocks that can occur 
in threaded programs when multiple threads access shared resources.

- Fault tolerance: If a process crashes in a multiprocessing program, it does not affect 
other processes, which can continue running. In a threaded program, a crash in one thread 
can bring down the entire process.

- Scalability: Multiprocessing can be more scalable than threading, as it can scale 
across multiple machines if needed, while threading is limited to a single machine.

Overall, while threading can be useful for simple concurrent tasks, multiprocessing is a 
more powerful and flexible approach for larger, more complex applications that require 
true parallelism and scalability.
"""
