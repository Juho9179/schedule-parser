# Schedule-parser

# What
Schedule-parser is a simple Python script that parses a workshift schedule which is in .xlsx format and has multiple employees in same file.

Schedule-parser then parses given employee's shifts and lists them in Google Calendar compatible format in CSV.
## WARNING: OVERWRITES old file, should it exist.

# How to run
```
python3 .\schedule-parser.py EMPLOYEE-NAME '.\schedule.xlsx'
```
Scheduler-parser then creates a new file, in the same directory as .xlsx file with same name, but .csv extension.


# Example
dummy_data.xlsx
![alt text](/docs/material.png)

```
python3 .\schedule-parser.py AB18 '.\example\dummy_data.xlsx'
-> outputs AB18's shifts in .\example\dummy_data.xlsx.csv
```

dummy_data.xlsx.csv

![alt text](/docs/result.png)

Scheduler-parser then creates a new file, in the same directory as .xlsx file with same name, but .csv extension.

# Disclaimer
THIS SOFTWARE IS PROVIDED BY THE OPEN CONNECTIVITY FOUNDATION, INC. "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE OR WARRANTIES OF NON-INFRINGEMENT, ARE DISCLAIMED. IN NO EVENT SHALL THE OPEN CONNECTIVITY FOUNDATION, INC. OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.