# Medical predictions (diabetis case)
This application converts a sequence of numbers into states. The states are arranged in a transition matrix and the transition probabilities are calculated for each element. The transition matrix is further used for this prediction in a Markov chain. For example, the application takes the following sequence of numbers:
```
159,82,187,194,179,115,197,102,105,104,95,126,74,143,143,127,98,70,92,170,168,182,149,85,137,100,170,180,61,177,86,195,198,182,150,197,103,103,186,100,96,196
```

The above sequence represents the glycemic values from a single individual who does not have diabetes, but a family predisposition for diabetes. Each number in the sequence represents a day. Thus, the sequence contains observations that extend over 42 days.

# Screenshot

<kbd><img src="https://github.com/Gagniuc/Diabetes-prediction-using-Markov-Chains/blob/main/screenshot/Medical%20prediction%20on%20diabetes.gif" /></kbd>

<kbd><img src="https://github.com/Gagniuc/Diabetes-prediction-using-Markov-Chains/blob/main/screenshot/How%20to%201.PNG" /></kbd>

<kbd><img src="https://github.com/Gagniuc/Diabetes-prediction-using-Markov-Chains/blob/main/screenshot/How%20to%202.PNG" /></kbd>

# References

<i>Paul A. Gagniuc. Markov chains: from theory to implementation and experimentation. Hoboken, NJ,  John Wiley & Sons, USA, 2017, ISBN: 978-1-119-38755-8.</i>
