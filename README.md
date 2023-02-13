# VBA-Epicor-MassSchedule
An Excel VBA macro to mass schedule a list of jobs backwards from an OP and Date. Some minor changes could be made to the code to have it schedule differently. Returns results of schedule in column explaining results of the scheduling.  Uses Epicor's Rest API.

The Excel sheet needs to be set up a specific way.

Row 1 is not looked at by the macro, they are for headers. Start data at Row 2

A is the column you put the jobs you want schedule in. 

B is for the op you are scheduling. Pick last op for the job if you are wanting to reschedule the entire job backwards. Op is by the number, not description.

C is for the date you are wanting to schedule to. Format is MM-DD-YYYY

Cell V9 needs to have the company name in it. Specifially how company name is referenced in epicor.

Column E will have any errors and the results of the scheduling pasted into it.
