import streamlit as st
import pandas as pd
import urllib.request
import matplotlib.pyplot as plt
from data import get_tourists, per_country, per_month, per_vehicle

st.title("ELSTAT Data analysis")

st.write("Download Data")


st.write("Total tourists")
fig = get_tourists()
st.pyplot(fig)

st.write("Total tourists by country")
fig = per_country()
st.pyplot(fig)

st.write("Total tourists by month")
fig = per_month()
st.pyplot(fig)

st.write("Total tourists by vehicle")
fig1, fig2, fig3, fig4 = per_vehicle()
st.pyplot(fig1)
st.pyplot(fig2)
st.pyplot(fig3)
st.pyplot(fig4)





