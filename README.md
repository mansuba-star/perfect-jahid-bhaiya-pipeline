# perfect-jahid-bhaiya-pipeline

#1st step to run this 

cd "/Users/mansuba/Desktop/main question validation /copy jahaid bhaiya project"

python data_formats.py data_tags_id_73.parquet data_tags_id_73_questions.txt


#2nd step to run this 
python detect_error.py \
  --questions data_tags_id_73_questions.txt \
  --reference data_tag_id_73_Stage4_final_output.txt \
  --output error_report_73.xlsx \
  --provider gemini \
  --api-key AIzaSyDLVXKzBFa3mIS1VxGoipe3JOxDyWYsQic
