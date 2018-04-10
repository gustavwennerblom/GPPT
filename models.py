from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import Column, Integer, String, Float, Binary

Base = declarative_base()


class Submission(Base):
    __tablename__ = 'gppt_submissions'
    Id = Column(Integer, primary_key=True)
    Filename = Column(String)
    Submitter = Column(String)
    Region = Column(String)
    Date = Column(String)
    Message_Id = Column(String)
    Attachment_Id = Column(String)
    Attachment_Binary = Column(Binary)
    Lead_Office = Column(String)
    P_Margin = Column(Float)
    Tot_Fee = Column(Integer)
    Blended_Rate = Column(Float)
    Tot_Hours = Column(Integer)
    Hours_Mgr = Column(Integer)
    Hours_SPM = Column(Integer)
    Hours_PM = Column(Integer)
    Hours_Cons = Column(Integer)
    Hours_Assoc = Column(Integer)
    Method = Column(String)
    Tool_Version = Column(String)
    SaleID = Column(Integer)
    ProjNo = Column(Integer)

    def __repr__(self):
        return "Submission Id {} from {} on {}. Tot_Fee: {}".format(self.Id, self.Submitter, self.Date, self.Tot_Fee)

class LastUpdated(Base):
    __tablename__ = 'last_update'
    Id = Column(Integer, primary_key=True)
    Updated = Column(String)

    def __repr__(self):
        return "Last updated timestamp with value {}".format(self.Updated)

class DevSubmission(Base):
    __tablename__ = 'dev_gppt_submissions'
    Id = Column(Integer, primary_key=True)
    Filename = Column(String)
    Submitter = Column(String)
    Region = Column(String)
    Date = Column(String)
    Message_Id = Column(String)
    Attachment_Id = Column(String)
    Attachment_Binary = Column(Binary)
    Lead_Office = Column(String)
    P_Margin = Column(Float)
    Tot_Fee = Column(Integer)
    Blended_Rate = Column(Float)
    Tot_Hours = Column(Integer)
    Hours_Mgr = Column(Integer)
    Hours_SPM = Column(Integer)
    Hours_PM = Column(Integer)
    Hours_Cons = Column(Integer)
    Hours_Assoc = Column(Integer)
    Method = Column(String)
    Tool_Version = Column(String)
    SaleID = Column(Integer)
    ProjNo = Column(Integer)

    def __repr__(self):
        return "(DEVELOPMENT) Submission Id {} from {} on {}. Tot_Fee: {}".format(self.Id, self.Submitter, self.Date, self.Tot_Fee)


class DevLastUpdated(Base):
    __tablename__ = 'dev_last_update'
    Id = Column(Integer, primary_key=True)
    Updated = Column(String)

    def __repr__(self):
        return "Last updated timestamp with value {}".format(self.Updated)
