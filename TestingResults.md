# A Comparison of VBA StringBuilder Implementations
The following implementations were tested
| Name | Implementation Notes |
| ---- | -------------------- |
| [StringChunker](https://gist.github.com/brucemcpherson/5102369) | String base, with in place modification |
| [StringBuffer](https://github.com/cristianbuse/VBA-StringBuffer) | String base, with in place modification, managed through a buffer |
| [clsStringBuilder](https://github.com/PeterRoach/VBA/blob/main/clsStringBuilder/clsStringBuilder.cls) | Byte array base |
| [StringBuilder](https://github.com/xformerfhs/VBAUtilities/blob/master/StringHandling/StringBuilder.cls) | String base, with in place modification, managed through a buffer |

**All times are in miliseconds and represent the average of 30 repetitions.**

## My Results

The first append test involved appending a single character to the end of the string using the library's append or concatenation function. 
| Iterations | Default | clsStringBuilder | StringChunker | StringBuffer | StringBuilder |
| :--------: | :-----: | :--------------: | :-----------: | :----------: |  :-----------: |
| 10<sup>2</sup> | 0 | 1 | 0 | 0 | 0 |
| 10<sup>3</sup> | 2 | 1 | 1 | 0 | 1 |
| 10<sup>4</sup> | 170 | 8 | 6 | 2 | 3 |
| 10<sup>5</sup> | N/A | 103 | 58 | 23 | 30 |
| 10<sup>6</sup> | N/A | 747 | 525 | 155 | 320 |

An additional test was performed to better simulate real word use by appending a word of random letters between 1 and 10 characters.
| Iterations | Default | clsStringBuilder | StringChunker | StringBuffer | StringBuilder |
| :--------: | :-----: | :--------------: | :-----------: | :----------: |  :-----------: |
| 10<sup>2</sup> | 0     | 1    | 1    | 1    | 1    |
| 10<sup>3</sup> | 3     | 4    | 3    | 3    | 3    |
| 10<sup>4</sup> | 81    | 32   | 25   | 21   | 22   |
| 10<sup>5</sup> | 31259 | 322  | 249  | 216  | 241  |
| 10<sup>6</sup> | N/A   | 3345 | 2551 | 2222 | 2386 |

## Further Reading
For more information on why the performance of the StringBuilder classes far outperform regular concatenation, refer to the following articles on the *Desktop Liberation* blog [here](https://ramblings.mcpher.com/optimization-links/strings-and-garbage/) and [here](https://ramblings.mcpher.com/strings-and-garbage-collector-in-vba/). 